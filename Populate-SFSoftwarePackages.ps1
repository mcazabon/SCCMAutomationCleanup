<#
.SYNOPSIS
    Populates SPDW_Staging.dbo.spdw_SFSoftwarePackages from a primary site database and sets EPP Status to Active.

.DESCRIPTION
    Reads packages (and optionally applications) from the primary site SCCM database, then upserts them
    into spdw_SFSoftwarePackages with EPPStatus = 'Active'. Items no longer in the site are flagged
    InSource = 0 (rows are never deleted). Uses Windows authentication.
    Prerequisite: run Deploy-EPPStatusSchema.sql first.

.PARAMETER SiteSqlServer
    SQL Server instance hosting the primary site database. Defaults to the UAT value in Library\Resources.xml.

.PARAMETER SiteDatabase
    Primary site database name (e.g. CM_US1). Defaults to the UAT value in Library\Resources.xml.

.PARAMETER SiteCode
    Site code recorded in SourceSite. Defaults to the UAT value in Library\Resources.xml.

.PARAMETER IncludeApplications
    Also load the latest version of each application.

.EXAMPLE
    .\Populate-SFSoftwarePackages.ps1 -WhatIf

.EXAMPLE
    .\Populate-SFSoftwarePackages.ps1 -SiteSqlServer 'SITESQL\INSTANCE' -SiteDatabase CM_XXX -SiteCode XXX
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$SiteSqlServer,
    [string]$SiteDatabase,
    [string]$SiteCode,
    [string]$SpdwSqlServer,
    [string]$SpdwDatabase,
    [string]$EppStatus = 'Active',
    [switch]$IncludeApplications
)

$ErrorActionPreference = 'Stop'

if (-not ($SiteSqlServer -and $SiteDatabase -and $SiteCode -and $SpdwSqlServer -and $SpdwDatabase)) {
    $xmlPath = Join-Path $PSScriptRoot 'Library\Resources.xml'
    [xml]$cfg = Get-Content -Path $xmlPath
    $uat = $cfg.sites.ChildNodes | Where-Object { $_.Attributes['uat'] } | Select-Object -First 1
    $spdw = $cfg.sites.ChildNodes | Where-Object { $_.Attributes['spdwsql'] } | Select-Object -First 1
    if (-not $SiteSqlServer) { $SiteSqlServer = $uat.sqlserver }
    if (-not $SiteDatabase)  { $SiteDatabase  = $uat.sccmsqldbname }
    if (-not $SiteCode)      { $SiteCode      = $uat.site }
    if (-not $SpdwSqlServer) { $SpdwSqlServer = $spdw.sqlserver }
    if (-not $SpdwDatabase)  { $SpdwDatabase  = $spdw.sqldbname }
}

function New-Connection([string]$Server, [string]$Database) {
    $conn = New-Object System.Data.SqlClient.SqlConnection(
        "Server=$Server; Database=$Database; Integrated Security=True; Connect Timeout=30; Application Name=Populate-SFSoftwarePackages")
    $conn.Open()
    return $conn
}

$packageSql = @"
SELECT N'Package' AS SourceType, CAST(PackageID AS NVARCHAR(256)) AS SourceID,
       CAST(Name AS NVARCHAR(500)) AS PackageName, CAST(NULL AS NVARCHAR(255)) AS PackageShortName
FROM dbo.v_Package WITH (NOLOCK)
WHERE PackageType = 0
"@

$applicationSql = @"
SELECT N'Application', CAST(CI_UniqueID AS NVARCHAR(256)),
       CAST(DisplayName AS NVARCHAR(500)), CAST(AdminComments AS NVARCHAR(255))
FROM dbo.fn_ListApplicationCIs(1033)
WHERE IsLatest = 1
"@

$sourceSql = $packageSql
if ($IncludeApplications) { $sourceSql += "`nUNION ALL`n$applicationSql" }

$siteConn = $null
$spdwConn = $null
try {
    Write-Host "Reading from $SiteSqlServer / $SiteDatabase (site $SiteCode)..."
    $siteConn = New-Connection $SiteSqlServer $SiteDatabase
    $cmd = $siteConn.CreateCommand()
    $cmd.CommandText = $sourceSql
    $cmd.CommandTimeout = 300
    $table = New-Object System.Data.DataTable
    [void](New-Object System.Data.SqlClient.SqlDataAdapter($cmd)).Fill($table)
    Write-Host "Retrieved $($table.Rows.Count) item(s)."

    # Packages have no comment column, so derive a short name: <ALPHANUMERIC NAME, max 24>_<PackageID>
    foreach ($row in $table.Rows) {
        if ($row['SourceType'] -eq 'Package' -and [string]::IsNullOrWhiteSpace([string]$row['PackageShortName'])) {
            $compact = ([string]$row['PackageName'] -replace '[^A-Za-z0-9]', '').ToUpperInvariant()
            if ($compact.Length -gt 24) { $compact = $compact.Substring(0, 24) }
            $row['PackageShortName'] = "${compact}_$($row['SourceID'])"
        }
    }

    if ($table.Rows.Count -eq 0) { throw "No items returned from $SiteDatabase - aborting to avoid flagging everything as removed." }

    if (-not $PSCmdlet.ShouldProcess("$SpdwSqlServer/$SpdwDatabase.dbo.spdw_SFSoftwarePackages", "Upsert $($table.Rows.Count) item(s) with EPPStatus '$EppStatus'")) { return }

    $spdwConn = New-Connection $SpdwSqlServer $SpdwDatabase
    $tran = $spdwConn.BeginTransaction()

    $create = $spdwConn.CreateCommand()
    $create.Transaction = $tran
    $create.CommandText = @"
CREATE TABLE #src (SourceType NVARCHAR(20) NOT NULL, SourceID NVARCHAR(256) NOT NULL,
                   PackageName NVARCHAR(500) NOT NULL, PackageShortName NVARCHAR(255) NULL,
                   PRIMARY KEY (SourceType, SourceID));
"@
    [void]$create.ExecuteNonQuery()

    $ins = $spdwConn.CreateCommand()
    $ins.Transaction = $tran
    $ins.CommandText = 'INSERT INTO #src (SourceType, SourceID, PackageName, PackageShortName) VALUES (@t, @i, @n, @s)'
    $pT = $ins.Parameters.Add('@t', [System.Data.SqlDbType]::NVarChar, 20)
    $pI = $ins.Parameters.Add('@i', [System.Data.SqlDbType]::NVarChar, 256)
    $pN = $ins.Parameters.Add('@n', [System.Data.SqlDbType]::NVarChar, 500)
    $pS = $ins.Parameters.Add('@s', [System.Data.SqlDbType]::NVarChar, 255)
    foreach ($row in $table.Rows) {
        $pT.Value = $row[0]; $pI.Value = $row[1]; $pN.Value = $row[2]; $pS.Value = $row[3]
        [void]$ins.ExecuteNonQuery()
    }

    $merge = $spdwConn.CreateCommand()
    $merge.Transaction = $tran
    $merge.CommandTimeout = 300
    $merge.CommandText = @"
UPDATE t SET PackageName = s.PackageName, PackageShortName = s.PackageShortName,
             EPPStatus = @EppStatus, InSource = 1, LastSynced = GETDATE()
FROM dbo.spdw_SFSoftwarePackages t
INNER JOIN #src s ON t.SourceType = s.SourceType AND t.SourceID = s.SourceID;
DECLARE @updated INT = @@ROWCOUNT;

INSERT INTO dbo.spdw_SFSoftwarePackages (SourceType, SourceID, PackageName, PackageShortName, EPPStatus)
SELECT s.SourceType, s.SourceID, s.PackageName, s.PackageShortName, @EppStatus
FROM #src s
WHERE NOT EXISTS (SELECT 1 FROM dbo.spdw_SFSoftwarePackages t WHERE t.SourceType = s.SourceType AND t.SourceID = s.SourceID);
DECLARE @inserted INT = @@ROWCOUNT;

UPDATE t SET InSource = 0, LastSynced = GETDATE()
FROM dbo.spdw_SFSoftwarePackages t
WHERE t.InSource = 1
  AND t.SourceType IN (SELECT SourceType FROM #src GROUP BY SourceType)
  AND NOT EXISTS (SELECT 1 FROM #src s WHERE s.SourceType = t.SourceType AND s.SourceID = t.SourceID);
DECLARE @removed INT = @@ROWCOUNT;

SELECT @inserted AS Inserted, @updated AS Updated, @removed AS Removed;
"@
    [void]$merge.Parameters.AddWithValue('@EppStatus', $EppStatus)

    $reader = $merge.ExecuteReader()
    [void]$reader.Read()
    $summary = "Inserted: {0}, updated: {1}, flagged removed: {2}" -f $reader['Inserted'], $reader['Updated'], $reader['Removed']
    $reader.Close()

    $tran.Commit()
    Write-Host "Done. $summary"
}
catch {
    if ($tran) { try { $tran.Rollback() } catch { } }
    Write-Error "Populate failed: $($_.Exception.Message)"
    exit 1
}
finally {
    if ($siteConn) { $siteConn.Dispose() }
    if ($spdwConn) { $spdwConn.Dispose() }
}
