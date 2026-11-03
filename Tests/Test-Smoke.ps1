<#
.SYNOPSIS
    Fast pre-flight smoke test for the SCCM cleanup scan (no scan is run unless -RunScan is given).
.DESCRIPTION
    Checks that scripts parse, the class module loads, Resources.xml is valid, the databases are
    reachable, the required tables/views exist, and the last scan left sane data.
    Exits with the number of failed checks.
.PARAMETER Server
    SQL Server instance hosting the databases. Defaults to the monitoring server in Library\Resources.xml.
.PARAMETER MonitoringDb
    Monitoring database name. Defaults to Library\Resources.xml.
.PARAMETER SccmDb
    SCCM database name. Defaults to the UAT site in Library\Resources.xml.
.PARAMETER StagingDb
    SPDW staging database name. Defaults to Library\Resources.xml.
.PARAMETER RunScan
    Also run a full scan in a fresh process and fail on any ERROR/FATAL/WARNING it logs.
#>
[CmdletBinding()]
param(
    [string]$Server,
    [string]$MonitoringDb,
    [string]$SccmDb,
    [string]$StagingDb,
    [switch]$RunScan
)

$root = Split-Path $PSScriptRoot -Parent
$script:Failures = 0

if (-not ($Server -and $MonitoringDb -and $SccmDb -and $StagingDb)) {
    [xml]$cfg = Get-Content (Join-Path $root 'Library\Resources.xml') -Raw
    $nodes = @($cfg.sites.ChildNodes)
    $uat = $nodes | Where-Object { $_.Attributes['uat'] } | Select-Object -First 1
    $spdw = $nodes | Where-Object { $_.Attributes['spdwsql'] } | Select-Object -First 1
    $mon = $nodes | Where-Object { $_.Attributes['monsql'] } | Select-Object -First 1
    if (-not $Server)       { $Server = $mon.sqlserver }
    if (-not $MonitoringDb) { $MonitoringDb = $mon.sqldbname }
    if (-not $SccmDb)       { $SccmDb = $uat.sccmsqldbname }
    if (-not $StagingDb)    { $StagingDb = $spdw.sqldbname }
}

function Check([string]$Name, [scriptblock]$Test) {
    try {
        $result = & $Test
        if ($result -eq $false) { throw "check returned false" }
        Write-Host "PASS  $Name"
    }
    catch {
        $script:Failures++
        Write-Host "FAIL  $Name -> $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Invoke-Scalar([string]$Db, [string]$Sql) {
    $conn = New-Object System.Data.SqlClient.SqlConnection "Server=$Server;Database=$Db;Integrated Security=True;Connection Timeout=15"
    try {
        $conn.Open()
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $Sql
        $cmd.CommandTimeout = 60
        return $cmd.ExecuteScalar()
    }
    finally { $conn.Dispose() }
}

# Files
foreach ($f in 'SCCMCollectionScan.ps1', 'Library\Classes.psm1', 'Library\Functions.ps1') {
    Check "Parses cleanly: $f" {
        $errs = $null; $tokens = $null
        [void][System.Management.Automation.Language.Parser]::ParseFile((Join-Path $root $f), [ref]$tokens, [ref]$errs)
        if ($errs.Count) { throw $errs[0].Message }
    }
}

Check 'Resources.xml is valid XML with a uat environment' {
    $xml = [xml](Get-Content (Join-Path $root 'Library\Resources.xml') -Raw)
    if (-not $xml.SelectSingleNode("//*[contains(local-name(),'uat')]") -and $xml.OuterXml -notmatch 'uat') { throw 'no uat section' }
}

# Class module, exercised in a child process because class types can't be re-imported in one session
Check 'Classes.psm1 loads and classes instantiate' {
    $tmp = Join-Path $env:TEMP 'smoke_classes.log'
    $code = "using module '$root\Library\Classes.psm1'; `$l=[Logger]::new('$tmp','uat'); `$e=[ErrorTracker]::new(); `$c=[ScriptConfig]::new('$root','uat',30); `$r=[ResourceManager]::new(`$l); `$l.Flush(); 'OK'"
    $out = powershell -NoProfile -ExecutionPolicy Bypass -Command $code 2>&1 | Out-String
    Remove-Item $tmp -ErrorAction SilentlyContinue
    if ($out -notmatch 'OK') { throw $out.Trim() }
}

Check 'Get-Help works on the main script' {
    $h = Get-Help (Join-Path $root 'SCCMCollectionScan.ps1')
    if (-not $h.Synopsis -or $h.Synopsis -match 'SCCMCollectionScan\.ps1') { throw 'help block not parsed' }
}

# Databases
foreach ($db in $SccmDb, $StagingDb, $MonitoringDb) {
    Check "Database reachable: $db" { Invoke-Scalar $db 'SELECT 1' | Out-Null }
}

Check 'SCCM objects exist (Folders, v_Collections)' {
    if ((Invoke-Scalar $SccmDb "SELECT COUNT(*) FROM sys.objects WHERE name IN ('Folders','v_Collections')") -ne 2) { throw 'missing' }
}
Check 'Staging EPP view has rows' {
    if ((Invoke-Scalar $StagingDb 'SELECT COUNT(*) FROM dbo.spdw_vwSFSoftwarePackages') -lt 1) { throw 'empty or missing' }
}
Check 'Monitoring tables and views exist' {
    $n = Invoke-Scalar $MonitoringDb "SELECT COUNT(*) FROM sys.objects WHERE name IN ('tbl_SCCMSites','tbl_SCCMCollections','tbl_SCCMapplications','tbl_SCCMtasksequences','tbl_SCCMCollectionStatus','tbl_SCCMTaskSequencesEPPCleanup','vw_EPPCleanup','vw_CollectionObjectsScan','vw_EPPConflicts')"
    if ($n -ne 9) { throw "found $n of 9" }
}
Check 'Monitoring views are queryable' {
    foreach ($v in 'vw_EPPCleanup', 'vw_CollectionObjectsScan', 'vw_EPPConflicts') { Invoke-Scalar $MonitoringDb "SELECT COUNT(*) FROM dbo.$v" | Out-Null }
}

if ($RunScan) {
    $start = Get-Date
    Check 'Full scan completes without ERROR/FATAL/WARNING' {
        $p = Start-Process powershell -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', (Join-Path $root 'SCCMCollectionScan.ps1'), '-environment', 'uat', '-days', '365' -Wait -PassThru -WindowStyle Hidden
        if ($p.ExitCode -ne 0) { throw "exit code $($p.ExitCode)" }
        $log = Get-ChildItem (Join-Path $root 'logs') -Filter 'SCCMCollectionScan_UAT_*.log' | Sort-Object LastWriteTime | Select-Object -Last 1
        $stamp = $start.ToString('yyyy-MM-dd HH:mm:ss')
        $bad = Get-Content $log.FullName | Where-Object { $_ -match '^\d{4}-' -and $_.Substring(0, 19) -ge $stamp -and $_ -match '\|(ERROR|FATAL|WARNING)\|' }
        if ($bad) { throw "$(@($bad).Count) problem lines, first: $(@($bad)[0])" }
    }
}

# Data sanity (meaningful after a scan has run)
Check 'Scan data present: collections and EPP cleanup rows' {
    if ((Invoke-Scalar $MonitoringDb 'SELECT COUNT(*) FROM dbo.tbl_SCCMCollections') -lt 1) { throw 'no collections' }
    if ((Invoke-Scalar $MonitoringDb 'SELECT COUNT(*) FROM dbo.tbl_SCCMTaskSequencesEPPCleanup') -lt 1) { throw 'no EPP cleanup rows' }
}
Check 'No doubled apostrophes in stored names' {
    $n = Invoke-Scalar $MonitoringDb "SELECT (SELECT COUNT(*) FROM dbo.tbl_SCCMCollections WHERE name LIKE '%''''%') + (SELECT COUNT(*) FROM dbo.tbl_SCCMtasksequences WHERE name LIKE '%''''%') + (SELECT COUNT(*) FROM dbo.tbl_SCCMapplications WHERE name LIKE '%''''%')"
    if ($n -ne 0) { throw "$n rows" }
}
Check 'Task sequence/application collection IDs exist in SCCM (or are the -1 fallback)' {
    foreach ($t in 'tbl_SCCMtasksequences', 'tbl_SCCMapplications') {
        $n = Invoke-Scalar $MonitoringDb "SELECT COUNT(*) FROM dbo.$t x WHERE x.collectionid <> '-1' AND NOT EXISTS (SELECT 1 FROM [$SccmDb].dbo.v_Collections c WHERE CAST(c.CollectionID AS nvarchar(50)) = x.collectionid)"
        if ($n -ne 0) { throw "$n bad rows in $t" }
    }
}
Check 'EPP conflicts view only shows Live/Pending Review/unverified statuses' {
    $n = Invoke-Scalar $MonitoringDb "SELECT COUNT(*) FROM dbo.vw_EPPConflicts WHERE ConflictType NOT IN ('Active package','Unverified package')"
    if ($n -ne 0) { throw "$n rows" }
}

Write-Host ""
Write-Host "Failures: $script:Failures"
exit $script:Failures
