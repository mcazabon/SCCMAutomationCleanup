function Get-ADObject {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$Filter,

        [Parameter(Mandatory = $false)]
        [string]$SearchBase,

        [Parameter(Mandatory = $false)]
        [string[]]$Properties = @('*')
    )

    # ADD: Check if DirectoryServices is available
    if (-not ([System.Management.Automation.PSTypeName]'System.DirectoryServices.DirectorySearcher').Type) {
        Write-Warning "System.DirectoryServices not available"
        return $null
    }

    try {
        # Create directory searcher
        $searcher = New-Object System.DirectoryServices.DirectorySearcher

        # Set search base if provided
        if ($SearchBase) {
            $entry = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$SearchBase")
            $searcher.SearchRoot = $entry
        }

        # Set filter (convert PowerShell filter to LDAP filter)
        if ($Filter) {
            # Simple conversion - you may need to enhance this
            $ldapFilter = $Filter -replace '-eq', '=' -replace '\$\.', ''
            $searcher.Filter = "($ldapFilter)"
        }
        else {
            $searcher.Filter = "(objectClass=*)"
        }

        # Add properties to load
        if ($Properties -ne @('*')) {
            foreach ($prop in $Properties) {
                $searcher.PropertiesToLoad.Add($prop) | Out-Null
            }
        }

        # Execute search
        $results = $searcher.FindAll()

        # Return results as custom objects
        foreach ($result in $results) {
            $obj = New-Object PSObject
            foreach ($propName in $result.Properties.PropertyNames) {
                $value = $result.Properties[$propName]
                if ($value.Count -eq 1) {
                    $obj | Add-Member -MemberType NoteProperty -Name $propName -Value $value[0]
                }
                else {
                    $obj | Add-Member -MemberType NoteProperty -Name $propName -Value $value
                }
            }
            Write-Output $obj
        }
    }
    catch {
        Write-Error "Error querying Active Directory: $_"
    }
    finally {
        if ($searcher) { $searcher.Dispose() }
        if ($entry) { $entry.Dispose() }
    }
}

function Get-ADGroup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Identity,

        [Parameter(Mandatory = $false)]
        [string[]]$Properties = @('whenChanged')
    )

    # ADD: Check if DirectoryServices is available
    if (-not ([System.Management.Automation.PSTypeName]'System.DirectoryServices.DirectorySearcher').Type) {
        Write-Warning "System.DirectoryServices not available"
        return $null
    }

    try {
        # Create directory searcher
        $searcher = New-Object System.DirectoryServices.DirectorySearcher
        $searcher.Filter = "(&(objectClass=group)(sAMAccountName=$Identity))"

        # Add properties
        foreach ($prop in $Properties) {
            $searcher.PropertiesToLoad.Add($prop) | Out-Null
        }

        # Execute search
        $result = $searcher.FindOne()
        if ($result) {
            $obj = New-Object PSObject
            foreach ($propName in $result.Properties.PropertyNames) {
                $value = $result.Properties[$propName]
                if ($value.Count -eq 1) {
                    $obj | Add-Member -MemberType NoteProperty -Name $propName -Value $value[0]
                }
                else {
                    $obj | Add-Member -MemberType NoteProperty -Name $propName -Value $value
                }
            }
            Write-Output $obj
        }
        else {
            Write-Error "Group '$Identity' not found"
        }
    }
    catch {
        Write-Error "Error querying Active Directory: $_"
    }
    finally {
        if ($searcher) { $searcher.Dispose() }
    }
}

function ConvertTo-SafeSqlString {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$Value,

        [Parameter(Mandatory = $false)]
        [string]$DefaultValue = ""
    )

    try {
        if ($null -eq $Value) {
            return $DefaultValue
        }

        # Values are always bound as SQL parameters, which handle quoting themselves.
        # Escaping here would store doubled quotes (O''Brien) and break duplicate checks.
        return $Value.ToString()
    }
    catch {
        return $DefaultValue
    }
}

function Build-EPPLookupTables {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Data.DataTable]$EPPData,

        [Parameter(Mandatory = $true)]
        $Logger
    )

    try {
        $Logger.Info("Building EPP lookup hashtables", "EPPLookup")

        # Hashtable keys are case-insensitive. When a name appears more than once, the row with
        # the most restrictive status wins (active beats unknown beats retired), so a Retired
        # duplicate can never hide a Live package.
        $eppByShortName = @{}
        $eppByPackageName = @{}
        $packageRows = New-Object 'System.Collections.Generic.List[object]'
        $duplicates = 0

        foreach ($row in $EPPData.Rows) {
            $shortName = "$($row["Package Short Name"])".Trim()
            $packageName = "$($row["Package Name"])".Trim()
            $rank = Get-EPPStatusRank $row["EPP Status"]

            foreach ($pair in @(
                @{ Table = $eppByShortName; Key = $shortName },
                @{ Table = $eppByPackageName; Key = $packageName }
            )) {
                if ($pair.Key.Length -eq 0) { continue }
                if ($pair.Table.ContainsKey($pair.Key)) {
                    $duplicates++
                    if ($rank -gt (Get-EPPStatusRank $pair.Table[$pair.Key]["EPP Status"])) {
                        $pair.Table[$pair.Key] = $row
                    }
                }
                else {
                    $pair.Table[$pair.Key] = $row
                }
            }

            if ($packageName.Length -gt 0) {
                $packageRows.Add(@{ Name = $packageName; Row = $row })
            }
        }

        $Logger.Info("EPP lookup tables built. Short names: $($eppByShortName.Count), Package names: $($eppByPackageName.Count), Duplicates resolved: $duplicates", "EPPLookup")
        return @{
            ByShortName = $eppByShortName
            ByPackageName = $eppByPackageName
            PackageRows = $packageRows
            Cache = @{}
        }
    }
    catch {
        $Logger.Error("Failed to build EPP lookup tables: $_", "EPPLookup")
        return @{
            ByShortName = @{}
            ByPackageName = @{}
            PackageRows = (New-Object 'System.Collections.Generic.List[object]')
            Cache = @{}
        }
    }
}

# Active (Live, Pending Review, ...) = 3, unknown/blank = 2, Retired/Cancel Package = 1
function Get-EPPStatusRank {
    param([AllowNull()]$Status)

    $text = "$Status".Trim()
    if ($text.Length -eq 0) { return 2 }
    if ($text -eq "Retired" -or $text -eq "Cancel Package") { return 1 }
    return 3
}

# Builds the lookup once per EPP table instead of once per call
function Get-EPPLookup {
    param(
        [System.Data.DataTable]$EPPData,
        $Logger
    )

    if ($null -ne $script:EPPLookupCache -and
        [object]::ReferenceEquals($script:EPPLookupCache.Source, $EPPData) -and
        $script:EPPLookupCache.RowCount -eq $EPPData.Rows.Count) {
        return $script:EPPLookupCache.Lookup
    }

    $lookup = Build-EPPLookupTables -EPPData $EPPData -Logger $Logger
    $script:EPPLookupCache = @{
        Source = $EPPData
        RowCount = $EPPData.Rows.Count
        Lookup = $lookup
    }
    return $lookup
}

# The single matcher used by collections, retired folders and task sequences.
# Order: short name from comments, short name from name, package name exact, package name contains.
function Find-EPPMatch {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Lookup,

        [AllowNull()]
        [string]$PackageName,

        [AllowNull()]
        [string]$AppComments
    )

    $name = "$PackageName".Trim()
    $comments = "$AppComments".Trim()
    $cacheKey = "$comments|$name"

    if ($Lookup.Cache.ContainsKey($cacheKey)) {
        return $Lookup.Cache[$cacheKey]
    }

    $matchedRow = $null
    $matchType = "Not Found"

    if ($comments.Length -gt 0 -and $Lookup.ByShortName.ContainsKey($comments)) {
        $matchedRow = $Lookup.ByShortName[$comments]
        $matchType = "Short Name (Exact)"
    }
    elseif ($name.Length -gt 0) {
        if ($Lookup.ByShortName.ContainsKey($name)) {
            $matchedRow = $Lookup.ByShortName[$name]
            $matchType = "Short Name (Exact)"
        }
        elseif ($Lookup.ByPackageName.ContainsKey($name)) {
            $matchedRow = $Lookup.ByPackageName[$name]
            $matchType = "Package Name (Exact)"
        }
        else {
            # Literal, case-insensitive substring (no wildcard parsing, so [ ] in names are safe).
            # Several hits resolve to the most restrictive status.
            foreach ($entry in $Lookup.PackageRows) {
                if ($entry.Name.IndexOf($name, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                    if ($null -eq $matchedRow -or
                        (Get-EPPStatusRank $entry.Row["EPP Status"]) -gt (Get-EPPStatusRank $matchedRow["EPP Status"])) {
                        $matchedRow = $entry.Row
                    }
                }
            }
            if ($null -ne $matchedRow) { $matchType = "Package Name (Contains)" }
        }
    }

    $result = @{ Row = $matchedRow; MatchType = $matchType }
    $Lookup.Cache[$cacheKey] = $result
    return $result
}

function Get-EPPPackageStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Data.DataTable]$EPPData,

        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [AllowNull()]
        [string]$PackageName,

        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [AllowNull()]
        [string]$AppComments,

        [Parameter(Mandatory = $true)]
        $Logger
    )

    try {
        # Validate input - return early if no valid package name
        if ([string]::IsNullOrWhiteSpace($PackageName) -or $PackageName -eq "N/A") {
            $Logger.Debug("No package name provided - skipping EPP check", "EPPMatcher")
            return @{
                Status = "Not Found"
                Bypass = $true
                MatchType = "No Package Name"
            }
        }

        $Logger.Debug("Checking EPP status for package: $PackageName", "EPPMatcher")

        # An empty EPP table means the EPP load failed; treat as unverifiable, not as "not found"
        if ($EPPData.Rows.Count -eq 0) {
            return @{
                Status = "Error"
                Bypass = $true
                MatchType = "EPP Data Unavailable"
            }
        }
        $lookup = Get-EPPLookup -EPPData $EPPData -Logger $Logger
        $match = Find-EPPMatch -Lookup $lookup -PackageName $PackageName -AppComments $AppComments
        $matchedRow = $match.Row
        $matchType = $match.MatchType

        if ($matchedRow) {
            $Logger.Debug("EPP match found via $matchType for package: $PackageName", "EPPMatcher")
        }

        # Process matched row
        if ($matchedRow) {
            $eppStatus = $matchedRow."EPP Status"
            if ([string]::IsNullOrWhiteSpace($eppStatus)) {
                $statusResult = "Null"
                $bypass = $true
            }
            elseif ($eppStatus -eq "Retired" -or $eppStatus -eq "Cancel Package") {
                $statusResult = $eppStatus
                $bypass = $false
            }
            else {
                $statusResult = $eppStatus
                $bypass = $true
            }

            $Logger.Debug("EPP match found - Status: $statusResult, Type: $matchType, Bypass: $bypass", "EPPMatcher")
            return @{
                Status = $statusResult
                Bypass = $bypass
                MatchType = $matchType
            }
        }
        else {
            $Logger.Debug("No EPP match found for package: $PackageName", "EPPMatcher")
            return @{
                Status = "Not Found"
                Bypass = $true
                MatchType = "Not Found"
            }
        }
    }
    catch {
        $Logger.Warning("Error matching EPP package status: $_", "EPPMatcher")
        return @{
            Status = "Error"
            Bypass = $true
            MatchType = "Error"
        }
    }
}

function Get-TaskSequenceEPPMatches {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Data.DataTable]$TaskSequenceData,

        [Parameter(Mandatory = $true)]
        [hashtable]$Lookup,

        [Parameter(Mandatory = $true)]
        $Logger
    )

    try {
        $Logger.Info("Matching Task Sequences with EPP data", "TSEPPMatch")

        $results = foreach ($row in $TaskSequenceData.Rows) {
            $refName = $row["Referenced Package Name"]
            if ($refName -is [DBNull] -or [string]::IsNullOrWhiteSpace($refName)) { continue }

            $match = Find-EPPMatch -Lookup $Lookup -PackageName $refName.ToString() -AppComments ""
            $eppMatch = $match.Row
            $matchType = $match.MatchType
            if ($eppMatch) {
                [PSCustomObject]@{
                    "Task Sequence Name" = $row["Task Sequence Name"]
                    "Task Sequence Package ID" = $row["Task Sequence Package ID"]
                    "Task Sequence Object Path" = $row["Task Sequence Object Path"]
                    "Referenced Package Name" = $refName
                    "ReferencePackageID" = $row["ReferencePackageID"]
                    "Collection Name" = $row["Collection Name"]
                    "Collection ID" = $row["CollectionID"]
                    "Collection Type" = $row["Collection Type"]
                    "Collection Object Path" = $row["Collection Object Path"]
                    "EPP Status" = $eppMatch["EPP Status"]
                    "Match Type" = $matchType
                }
            }
        }

        $Logger.Info("Task Sequence EPP matching completed. Matches found: $($results.Count)", "TSEPPMatch")
        return $results
    }
    catch {
        $Logger.Error("Failed to match Task Sequences with EPP: $_", "TSEPPMatch")
        return @()
    }
}

function Get-ADGroupInformation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$CollectionName,

        [Parameter(Mandatory = $true)]
        [string]$CollectionType,

        [Parameter(Mandatory = $true)]
        [datetime]$ReferenceDate,

        [Parameter(Mandatory = $true)]
        $Logger,

        [Parameter(Mandatory = $true)]
        $ErrorTracker
    )

    try {
        $result = @{
            Exists = $false
            WhenChanged = $null
            WhenChangedString = $null
            DaysSinceChange = 0
        }

        # Only check for user collections starting with 'u' or 'U'
        if ($CollectionType -ne "UserCollection") {
            $Logger.Debug("Skipping AD check - not a user collection", "ADGroupCheck")
            return $result
        }

        if (-not ($CollectionName.StartsWith("u") -or $CollectionName.StartsWith("U"))) {
            $Logger.Debug("Skipping AD check - collection name doesn't start with 'u'", "ADGroupCheck")
            return $result
        }

        # FIX: Check if AD cmdlets are available
        if (-not (Get-Command Get-ADObject -ErrorAction SilentlyContinue)) {
            $Logger.Warning("Get-ADObject cmdlet not available - skipping AD check", "ADGroupCheck")
            return $result
        }

        $adGroupName = $CollectionName.Substring(1)
        $Logger.Debug("Checking AD group: $adGroupName", "ADGroupCheck")

        try {
            # FIX: Use proper error handling
            $adObject = Get-ADObject -Filter "sAMAccountName -eq '$adGroupName'" -ErrorAction Stop
            if ($null -ne $adObject) {
                # FIX: Check if Get-ADGroup is available
                if (-not (Get-Command Get-ADGroup -ErrorAction SilentlyContinue)) {
                    $Logger.Warning("Get-ADGroup cmdlet not available", "ADGroupCheck")
                    return $result
                }

                $adGroup = Get-ADGroup -Identity $adGroupName -Properties whenChanged -ErrorAction Stop
                $result.Exists = $true
                $result.WhenChanged = $adGroup.whenChanged
                $result.WhenChangedString = Get-Date $adGroup.whenChanged -Format "yyyy-MM-dd HH:mm:ss"
                $result.DaysSinceChange = ($ReferenceDate - $adGroup.whenChanged).Days
                $Logger.Debug("AD group found - Last changed: $($result.WhenChangedString) ($($result.DaysSinceChange) days ago)", "ADGroupCheck")
            }
            else {
                $Logger.Debug("AD group not found: $adGroupName", "ADGroupCheck")
            }
        }
        catch {
            # FIX: Don't record as error - just log as debug
            $Logger.Debug("AD group check failed for '$adGroupName': $_", "ADGroupCheck")
        }

        return $result
    }
    catch {
        $Logger.Debug("AD group check failed: $_", "ADGroupCheck")
        return @{
            Exists = $false
            WhenChanged = $null
            WhenChangedString = $null
            DaysSinceChange = 0
        }
    }
}

function Show-LogFileSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LogDir
    )

    $logFiles = Get-LogFileList -LogDir $LogDir -DaysBack 30

    if ($logFiles.Count -eq 0) {
        Write-Host "No log files found in the last 30 days." -ForegroundColor Yellow
        return
    }

    Write-Host "`n=== Log File Summary ===" -ForegroundColor Cyan
    Write-Host "Total log files: $($logFiles.Count)" -ForegroundColor White
    Write-Host ""

    $logFiles | ForEach-Object {
        $errorCount = (Select-String -Path $_.FullName -Pattern "\|ERROR\|" -ErrorAction SilentlyContinue).Count
        $warningCount = (Select-String -Path $_.FullName -Pattern "\|WARNING\|" -ErrorAction SilentlyContinue).Count
        $fatalCount = (Select-String -Path $_.FullName -Pattern "\|FATAL\|" -ErrorAction SilentlyContinue).Count

        $color = if ($fatalCount -gt 0) { 'Red' }
        elseif ($errorCount -gt 0) { 'Yellow' }
        else { 'Green' }

        Write-Host "$($_.Name)" -ForegroundColor $color -NoNewline
        Write-Host " - Size: $([math]::Round($_.Length/1KB, 2)) KB" -ForegroundColor Gray -NoNewline
        Write-Host " - Errors: $errorCount, Warnings: $warningCount, Fatal: $fatalCount" -ForegroundColor Gray
    }

    Write-Host "`nMost recent log: $($logFiles[0].FullName)" -ForegroundColor Cyan
    Write-Host "========================`n" -ForegroundColor Cyan
}

function Get-LogFileList {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LogDir,

        [Parameter(Mandatory = $false)]
        [int]$DaysBack = 30
    )

    try {
        $cutoffDate = (Get-Date).AddDays(-$DaysBack)
        $logFiles = Get-ChildItem -Path $LogDir -Filter "SCCMCollectionScan_*.log" -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -ge $cutoffDate } |
            Sort-Object LastWriteTime -Descending

        return $logFiles
    }
    catch {
        Write-Warning "Failed to retrieve log file list: $_"
        return @()
    }
}

function Export-LogAnalysis {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LogFile,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    try {
        $logContent = Get-Content $LogFile -ErrorAction Stop

        $analysis = @{
            LogFile = $LogFile
            TotalLines = $logContent.Count
            InfoCount = ($logContent | Select-String -Pattern "\|INFO\|").Count
            WarningCount = ($logContent | Select-String -Pattern "\|WARNING\|").Count
            ErrorCount = ($logContent | Select-String -Pattern "\|ERROR\|").Count
            FatalCount = ($logContent | Select-String -Pattern "\|FATAL\|").Count
            DebugCount = ($logContent | Select-String -Pattern "\|DEBUG\|").Count
            SuccessCount = ($logContent | Select-String -Pattern "\|SUCCESS\|").Count
            Errors = $logContent | Select-String -Pattern "\|ERROR\|" | ForEach-Object { $_.Line }
            Warnings = $logContent | Select-String -Pattern "\|WARNING\|" | ForEach-Object { $_.Line }
            Fatals = $logContent | Select-String -Pattern "\|FATAL\|" | ForEach-Object { $_.Line }
        }

        $analysis | ConvertTo-Json -Depth 10 | Out-File $OutputPath -Encoding UTF8
        Write-Host "Log analysis exported to: $OutputPath" -ForegroundColor Green
        return $analysis
    }
    catch {
        Write-Error "Failed to export log analysis: $_"
        return $null
    }
}

function Select-RetiredTaskSequenceMatches {
    # A task sequence is a candidate only when every matched package is retired/cancelled;
    # one active or unverified package keeps it out.
    param(
        [object[]]$Matches,
        [object[]]$ExcludePackageIds = @()
    )

    $blocked = @{}
    foreach ($m in $Matches) {
        if ($m.'EPP Status' -notin @("Retired", "Cancel Package")) {
            $blocked[[string]$m.'Task Sequence Package ID'] = $true
        }
    }

    @($Matches | Where-Object {
        $id = [string]$_.'Task Sequence Package ID'
        $_.'EPP Status' -in @("Retired", "Cancel Package") -and
        -not $blocked.ContainsKey($id) -and
        $id -notin $ExcludePackageIds
    })
}

function Test-CollectionProcessingRules {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$CollectionItem,

        [Parameter(Mandatory = $true)]
        [hashtable]$EPPStatus,

        [Parameter(Mandatory = $true)]
        [hashtable]$ADGroupInfo,

        [Parameter(Mandatory = $true)]
        [datetime]$ReferenceDate,

        [Parameter(Mandatory = $true)]
        $Logger
    )

    try {
        $shouldProcess = $false
        $useADGroupDate = $false
        $reason = "Does not meet criteria"

        $collectionPath = $CollectionItem.CollectionPath
        $isStaging = $collectionPath -match "staging"
        $daysSinceChange = [int]$CollectionItem."Days Since Last Member Change"

        # Calculate threshold dates
        # -days sets the non-staging threshold; staging is 90 days or less when -days is lower
        $nonStagingDays = if ($script:NonStagingDays -gt 0) { [int]$script:NonStagingDays } else { 365 }
        $stagingDays = [Math]::Min(90, $nonStagingDays)
        $deleteDate365 = $ReferenceDate.AddDays(-$nonStagingDays)
        $deleteDate90 = $ReferenceDate.AddDays(-$stagingDays)

        # Exclusions: a collection that is still referenced must never be a candidate
        $exclusions = @()
        foreach ($flag in @(
            @{ Property = "LimitingCollection"; Label = "used as a limiting collection" },
            @{ Property = "ReferencedCollection"; Label = "referenced by another collection" },
            @{ Property = "App TS References"; Label = "application referenced by a task sequence" },
            @{ Property = "Adaptiva Client Settings Collection Reference"; Label = "Adaptiva client settings reference" },
            @{ Property = "Adaptiva Policy Collection Reference"; Label = "Adaptiva policy reference" },
            @{ Property = "Adaptiva Health Check Collection Reference"; Label = "Adaptiva health check reference" }
        )) {
            $value = $CollectionItem.($flag.Property)
            if ($null -ne $value -and $value -isnot [DBNull] -and "$value" -eq "True") {
                $exclusions += $flag.Label
            }
        }

        if ($exclusions.Count -gt 0) {
            $reason = "Excluded - " + ($exclusions -join "; ")
            $Logger.Debug("Rule skipped: $reason", "ProcessingRules")
            return @{
                ShouldProcess = $false
                UseADGroupDate = $false
                Reason = $reason
            }
        }

        # Age rules only apply when no active EPP package is involved. Live, Pending Review
        # and lookup errors must never let a collection age out.
        $ageEligible = $EPPStatus.Status -in @("Not Found", "Null")

        # Rule 1: EPP Status is Retired or Cancelled
        if ($EPPStatus.Status -eq "Retired" -or $EPPStatus.Status -eq "Cancel Package") {
            $shouldProcess = $true
            $reason = "EPP package is $($EPPStatus.Status)"
            $Logger.Debug("Rule matched: $reason", "ProcessingRules")
        }
        # Rule 2: Non-staging collection with AD group older than 365 days
        elseif ($ADGroupInfo.Exists -and
            $ADGroupInfo.WhenChanged -lt $deleteDate365 -and
            -not $isStaging -and
            $ageEligible) {
            $shouldProcess = $true
            $useADGroupDate = $true
            $reason = "Non-staging collection with AD group unchanged for $nonStagingDays+ days"
            $Logger.Debug("Rule matched: $reason", "ProcessingRules")
        }
        # Rule 3: Non-staging collection without AD group, older than 365 days
        elseif ($daysSinceChange -ge $nonStagingDays -and
            -not $isStaging -and
            -not $ADGroupInfo.Exists -and
            $ageEligible) {
            $shouldProcess = $true
            $reason = "Non-staging collection unchanged for $nonStagingDays+ days (no AD group)"
            $Logger.Debug("Rule matched: $reason", "ProcessingRules")
        }
        # Rule 4: Staging collection with AD group older than 90 days
        elseif ($ADGroupInfo.Exists -and
            $ADGroupInfo.WhenChanged -lt $deleteDate90 -and
            $isStaging -and
            $ageEligible) {
            $shouldProcess = $true
            $useADGroupDate = $true
            $reason = "Staging collection with AD group unchanged for $stagingDays+ days"
            $Logger.Debug("Rule matched: $reason", "ProcessingRules")
        }
        # Rule 5: Staging collection without AD group, older than 90 days
        elseif ($daysSinceChange -ge $stagingDays -and
            $isStaging -and
            -not $ADGroupInfo.Exists -and
            $ageEligible) {
            $shouldProcess = $true
            $reason = "Staging collection unchanged for $stagingDays+ days (no AD group)"
            $Logger.Debug("Rule matched: $reason", "ProcessingRules")
        }

        return @{
            ShouldProcess = $shouldProcess
            UseADGroupDate = $useADGroupDate
            Reason = $reason
        }
    }
    catch {
        $Logger.Error("Error evaluating processing rules: $_", "ProcessingRules")
        return @{
            ShouldProcess = $false
            UseADGroupDate = $false
            Reason = "Error during evaluation"
        }
    }
}

function Get-BaselineReferences {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $SCCMConnection,

        [Parameter(Mandatory = $true)]
        [string]$BaselineName,

        [Parameter(Mandatory = $true)]
        $Logger
    )

    try {
        $Logger.Debug("Checking baseline references for: $BaselineName", "BaselineRefChecker")

        $refQuery = @"
SET NOCOUNT ON

-- Check if baseline is deployed to collections
;WITH BaselineDeployments AS (
    SELECT DISTINCT
        CI.CI_ID,
        CI.CIType_ID,
        LCP.DisplayName AS BaselineName,
        C.Name AS CollectionName,
        'Active Deployment' AS ReferenceType
    FROM v_ConfigurationItems CI WITH (NOLOCK)
    INNER JOIN v_LocalizedCIProperties LCP WITH (NOLOCK)
        ON CI.CI_ID = LCP.CI_ID
    INNER JOIN v_CITypes CIT WITH (NOLOCK)
        ON CI.CIType_ID = CIT.CIType_ID
    LEFT JOIN v_CIAssignmentToCI CIACI WITH (NOLOCK)
        ON CIACI.CI_ID = CI.CI_ID
    LEFT JOIN v_CIAssignment CIA WITH (NOLOCK)
        ON CIACI.AssignmentID = CIA.AssignmentID
    LEFT JOIN v_Collection C WITH (NOLOCK)
        ON C.CollectionID = CIA.CollectionID
    WHERE CIT.TypeName = 'Baseline'
        AND LCP.DisplayName = @BaselineName
        AND C.Name IS NOT NULL
),
-- Check if baseline has configuration items
BaselineConfigItems AS (
    SELECT DISTINCT
        CI.CI_ID,
        LCP.DisplayName AS BaselineName,
        LCP_Child.DisplayName AS ConfigItemName,
        'Configuration Item Reference' AS ReferenceType
    FROM v_ConfigurationItems CI WITH (NOLOCK)
    INNER JOIN v_LocalizedCIProperties LCP WITH (NOLOCK)
        ON CI.CI_ID = LCP.CI_ID
    INNER JOIN v_CITypes CIT WITH (NOLOCK)
        ON CI.CIType_ID = CIT.CIType_ID
    LEFT JOIN v_CIRelation CIR WITH (NOLOCK)
        ON CI.CI_ID = CIR.FromCIID
    LEFT JOIN v_LocalizedCIProperties LCP_Child WITH (NOLOCK)
        ON CIR.ToCIID = LCP_Child.CI_ID
    WHERE CIT.TypeName = 'Baseline'
        AND LCP.DisplayName = @BaselineName
        AND LCP_Child.DisplayName IS NOT NULL
)
-- Combine all references
SELECT BaselineName, CollectionName AS ReferenceName, ReferenceType
FROM BaselineDeployments
UNION ALL
SELECT BaselineName, ConfigItemName AS ReferenceName, ReferenceType
FROM BaselineConfigItems
"@

        $dataset = New-Object System.Data.DataSet
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($refQuery, $SCCMConnection.Connection)
        $adapter.SelectCommand.CommandTimeout = 300

        # Add parameter
        [void]$adapter.SelectCommand.Parameters.AddWithValue("@BaselineName", $BaselineName)
        [void]$adapter.Fill($dataset)

        $references = $dataset.Tables[0]
        $isOrphaned = ($references.Rows.Count -eq 0)
        $Logger.Debug("Baseline '$BaselineName' - Orphaned: $isOrphaned, References: $($references.Rows.Count)", "BaselineRefChecker")

        return @{
            IsOrphaned = $isOrphaned
            ReferenceCount = $references.Rows.Count
            References = $references
        }
    }
    catch {
        $Logger.Warning("Failed to check baseline references for '$BaselineName': $_", "BaselineRefChecker")
        return @{
            IsOrphaned = $false
            ReferenceCount = -1
            References = $null
        }
    }
}

function Get-PackageReferences {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $SCCMConnection,

        [Parameter(Mandatory = $true)]
        [string]$PackageID,

        [Parameter(Mandatory = $true)]
        $Logger
    )

    try {
        $Logger.Debug("Checking package references for: $PackageID", "PackageRefChecker")

        # Removed DECLARE statement - use parameter directly
        $refQuery = @"
SET NOCOUNT ON

-- Check if package is referenced by task sequences
;WITH TSReferences AS (
    SELECT DISTINCT
        tsp.Name AS TaskSequenceName,
        pkg.Name AS PackageName,
        pkg.PackageID,
        'Task Sequence Reference' AS ReferenceType
    FROM v_TaskSequencePackage tsp
    JOIN v_TaskSequenceReferencesInfo tsri
        ON tsp.PackageID = tsri.PackageID
    JOIN v_Package pkg
        ON tsri.ReferencePackageID = pkg.PackageID
    WHERE pkg.PackageID = @PackageID
),
-- Check if package has active advertisements
AdvertisementReferences AS (
    SELECT DISTINCT
        adv.AdvertisementName,
        pkg.Name AS PackageName,
        pkg.PackageID,
        'Active Advertisement' AS ReferenceType
    FROM v_Package pkg
    INNER JOIN v_Advertisement adv
        ON pkg.PackageID = adv.PackageID
    WHERE pkg.PackageID = @PackageID
),
-- Check if package is in distribution point groups
DPGroupReferences AS (
    SELECT DISTINCT
        dpgr.Name AS DPGroupName,
        pkg.Name AS PackageName,
        pkg.PackageID,
        'Distribution Point Group' AS ReferenceType
    FROM v_Package pkg
    INNER JOIN v_dpgrouppackages dgp
        ON pkg.PackageID = dgp.pkgid
    INNER JOIN vsms_dpgroupinfo dpgr
        ON dgp.groupid = dpgr.groupid
    WHERE pkg.PackageID = @PackageID
)
-- Combine all references
SELECT TaskSequenceName AS ReferenceName, PackageName AS ObjectName, PackageID, ReferenceType
FROM TSReferences
UNION ALL
SELECT AdvertisementName AS ReferenceName, PackageName AS ObjectName, PackageID, ReferenceType
FROM AdvertisementReferences
UNION ALL
SELECT DPGroupName AS ReferenceName, PackageName AS ObjectName, PackageID, ReferenceType
FROM DPGroupReferences
"@

        $dataset = New-Object System.Data.DataSet
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($refQuery, $SCCMConnection.Connection)
        $adapter.SelectCommand.CommandTimeout = 300

        # Add parameter - this is the ONLY place @PackageID should be defined
        [void]$adapter.SelectCommand.Parameters.AddWithValue("@PackageID", $PackageID)
        [void]$adapter.Fill($dataset)

        $references = $dataset.Tables[0]
        $isOrphaned = ($references.Rows.Count -eq 0)
        $Logger.Debug("Package '$PackageID' - Orphaned: $isOrphaned, References: $($references.Rows.Count)", "PackageRefChecker")

        return @{
            IsOrphaned = $isOrphaned
            ReferenceCount = $references.Rows.Count
            References = $references
        }
    }
    catch {
        $Logger.Warning("Failed to check package references for '$PackageID': $_", "PackageRefChecker")
        return @{
            IsOrphaned = $false
            ReferenceCount = -1
            References = $null
        }
    }
}

function Get-TaskSequenceReferencedPackages {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $SCCMConnection,

        [Parameter(Mandatory = $true)]
        [string]$TaskSequenceID,

        [Parameter(Mandatory = $true)]
        $Logger
    )

    try {
        $Logger.Debug("Retrieving referenced packages for task sequence: $TaskSequenceID", "TSReferencedPackages")

        $packageQuery = @"
SET NOCOUNT ON
SELECT DISTINCT
    tsri.ReferencePackageID AS PackageID,
    pkg.Name AS PackageName,
    CASE pkg.PackageType
        WHEN 0 THEN 'Software Distribution Package'
        WHEN 3 THEN 'Driver Package'
        WHEN 5 THEN 'Software Update Package'
        WHEN 257 THEN 'Image Package'
        WHEN 258 THEN 'Boot Image Package'
        WHEN 259 THEN 'OS Install Package'
        ELSE 'Other'
    END AS PackageType
FROM v_TaskSequenceReferencesInfo tsri
INNER JOIN v_Package pkg
    ON tsri.ReferencePackageID = pkg.PackageID
WHERE tsri.PackageID = @TaskSequenceID
ORDER BY pkg.Name
"@

        $dataset = New-Object System.Data.DataSet
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($packageQuery, $SCCMConnection.Connection)
        $adapter.SelectCommand.CommandTimeout = 300

        # Add parameter
        [void]$adapter.SelectCommand.Parameters.AddWithValue("@TaskSequenceID", $TaskSequenceID)
        [void]$adapter.Fill($dataset)

        $packages = $dataset.Tables[0]
        $Logger.Debug("Found $($packages.Rows.Count) referenced packages for task sequence $TaskSequenceID", "TSReferencedPackages")
        return $packages.Rows
    }
    catch {
        $Logger.Warning("Failed to retrieve referenced packages for task sequence $TaskSequenceID : $_", "TSReferencedPackages")
        return @()
    }
}

function Get-CollectionAssociatedObjects {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $SCCMConnection,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$CollectionID,

        [Parameter(Mandatory = $true)]
        $Logger
    )

    $dataset = $null
    $adapter = $null
    $command = $null

    try {
        $Logger.Debug(
            "Retrieving associated objects for collection: $CollectionID",
            "AssociatedObjects"
        )

        $objectQuery = @"
SET NOCOUNT ON;
;WITH CollectionApplications AS (
    SELECT DISTINCT
        CAST('Application' AS NVARCHAR(50)) AS ObjectType,
        CAST(app.DisplayName AS NVARCHAR(512)) AS ObjectName,
        CAST(app.ModelName AS NVARCHAR(512)) AS ObjectID,
        CAST(app.ModelName AS NVARCHAR(512)) AS ModelName,
        CAST(dep.CollectionID AS NVARCHAR(50)) AS CollectionID,
        CAST(
            CASE
                WHEN dep.DeploymentIntent = 1 THEN 'Required'
                WHEN dep.DeploymentIntent = 2 THEN 'Available'
                ELSE 'Unknown'
            END AS NVARCHAR(50)
        ) AS DeploymentIntent
    FROM v_DeploymentSummary AS dep WITH (NOLOCK)
    INNER JOIN dbo.fn_ListApplicationCIs(1033) AS app
        ON CAST(dep.ModelId AS NVARCHAR(512)) = CAST(app.ModelName AS NVARCHAR(512))
    WHERE CAST(dep.CollectionID AS NVARCHAR(50)) = @CollectionID
        AND dep.FeatureType = 1
        AND app.IsLatest = 1
),
CollectionPackages AS (
    SELECT DISTINCT
        CAST('Package' AS NVARCHAR(50)) AS ObjectType,
        CAST(pkg.Name AS NVARCHAR(512)) AS ObjectName,
        CAST(pkg.PackageID AS NVARCHAR(512)) AS ObjectID,
        CAST(NULL AS NVARCHAR(512)) AS ModelName,
        CAST(adv.CollectionID AS NVARCHAR(50)) AS CollectionID,
        CAST(
            CASE
                WHEN (adv.AdvertFlags & 0x00000020) = 0x00000020 THEN 'Required'
                ELSE 'Available'
            END AS NVARCHAR(50)
        ) AS DeploymentIntent
    FROM v_Advertisement AS adv WITH (NOLOCK)
    INNER JOIN v_Package AS pkg WITH (NOLOCK)
        ON CAST(adv.PackageID AS NVARCHAR(512)) = CAST(pkg.PackageID AS NVARCHAR(512))
    WHERE CAST(adv.CollectionID AS NVARCHAR(50)) = @CollectionID
        AND pkg.PackageType NOT IN (4, 8)
),
CollectionBaselines AS (
    SELECT DISTINCT
        CAST('Baseline' AS NVARCHAR(50)) AS ObjectType,
        CAST(lci.DisplayName AS NVARCHAR(512)) AS ObjectName,
        CAST(vcf.ModelID AS NVARCHAR(512)) AS ObjectID,
        CAST(vcf.ModelID AS NVARCHAR(512)) AS ModelName,
        CAST(dep.CollectionID AS NVARCHAR(50)) AS CollectionID,
        CAST(
            CASE
                WHEN dep.DeploymentIntent = 1 THEN 'Required'
                WHEN dep.DeploymentIntent = 2 THEN 'Available'
                ELSE 'Monitor'
            END AS NVARCHAR(50)
        ) AS DeploymentIntent
    FROM v_DeploymentSummary AS dep WITH (NOLOCK)
    INNER JOIN v_ConfigurationItems AS vcf WITH (NOLOCK)
        ON CAST(dep.ModelId AS NVARCHAR(512)) = CAST(vcf.ModelID AS NVARCHAR(512))
    INNER JOIN v_CITypes AS cit WITH (NOLOCK)
        ON vcf.CIType_ID = cit.CIType_ID
    INNER JOIN v_LocalizedCIProperties AS lci WITH (NOLOCK)
        ON vcf.CI_ID = lci.CI_ID
    WHERE CAST(dep.CollectionID AS NVARCHAR(50)) = @CollectionID
        AND cit.TypeName = 'Baseline'
        AND dep.FeatureType = 6
),
CollectionTaskSequences AS (
    SELECT DISTINCT
        CAST('TaskSequence' AS NVARCHAR(50)) AS ObjectType,
        CAST(tsp.Name AS NVARCHAR(512)) AS ObjectName,
        CAST(tsp.PackageID AS NVARCHAR(512)) AS ObjectID,
        CAST(NULL AS NVARCHAR(512)) AS ModelName,
        CAST(adv.CollectionID AS NVARCHAR(50)) AS CollectionID,
        CAST(
            CASE
                WHEN (adv.AdvertFlags & 0x00000020) = 0x00000020 THEN 'Required'
                ELSE 'Available'
            END AS NVARCHAR(50)
        ) AS DeploymentIntent
    FROM v_Advertisement AS adv WITH (NOLOCK)
    INNER JOIN v_TaskSequencePackage AS tsp WITH (NOLOCK)
        ON CAST(adv.PackageID AS NVARCHAR(512)) = CAST(tsp.PackageID AS NVARCHAR(512))
    WHERE CAST(adv.CollectionID AS NVARCHAR(50)) = @CollectionID
)
SELECT ObjectType, ObjectName, ObjectID, ModelName, DeploymentIntent
FROM CollectionApplications
UNION ALL
SELECT ObjectType, ObjectName, ObjectID, ModelName, DeploymentIntent
FROM CollectionPackages
UNION ALL
SELECT ObjectType, ObjectName, ObjectID, ModelName, DeploymentIntent
FROM CollectionBaselines
UNION ALL
SELECT ObjectType, ObjectName, ObjectID, ModelName, DeploymentIntent
FROM CollectionTaskSequences
ORDER BY ObjectType, ObjectName;
"@

        if ($null -eq $SCCMConnection.Connection) {
            throw "SCCM SQL connection is null."
        }

        if ($SCCMConnection.Connection.State -ne [System.Data.ConnectionState]::Open) {
            $SCCMConnection.Connection.Open()
        }

        $dataset = New-Object System.Data.DataSet
        $command = $SCCMConnection.Connection.CreateCommand()
        $command.CommandText = $objectQuery
        $command.CommandTimeout = 300

        $parameter = $command.Parameters.Add(
            "@CollectionID",
            [System.Data.SqlDbType]::NVarChar,
            50
        )
        $parameter.Value = $CollectionID

        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $adapter.SelectCommand = $command
        [void]$adapter.Fill($dataset)

        if ($dataset.Tables.Count -eq 0) {
            $Logger.Debug(
                "Query returned no result table for collection $CollectionID",
                "AssociatedObjects"
            )
            return @()
        }

        $objects = $dataset.Tables[0]
        if ($objects.Rows.Count -eq 0) {
            $Logger.Debug(
                "No associated objects found for collection $CollectionID",
                "AssociatedObjects"
            )
            return @()
        }

        $appCount = @($objects.Select("ObjectType = 'Application'")).Count
        $pkgCount = @($objects.Select("ObjectType = 'Package'")).Count
        $baselineCount = @($objects.Select("ObjectType = 'Baseline'")).Count
        $tsCount = @($objects.Select("ObjectType = 'TaskSequence'")).Count

        $Logger.Debug(
            "Found $($objects.Rows.Count) associated objects for collection $CollectionID",
            "AssociatedObjects"
        )
        $Logger.Debug(
            "Breakdown - Apps: $appCount, Packages: $pkgCount, Baselines: $baselineCount, Task Sequences: $tsCount",
            "AssociatedObjects"
        )

        return @($objects.Rows)
    }
    catch {
        $errorMessage = $_.Exception.Message
        if ($_.Exception.InnerException) {
            $errorMessage += " | Inner Exception: $($_.Exception.InnerException.Message)"
        }

        $Logger.Warning(
            "Failed to retrieve associated objects for collection $CollectionID : $errorMessage",
            "AssociatedObjects"
        )
        return @()
    }
    finally {
        if ($null -ne $adapter) {
            $adapter.Dispose()
        }
        if ($null -ne $command) {
            $command.Dispose()
        }
        if ($null -ne $dataset) {
            $dataset.Dispose()
        }
    }
}
