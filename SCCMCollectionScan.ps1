<#
.SYNOPSIS
    SCCM Collection Scan and Cleanup Analysis Tool with EPP Integration, Retired Folder Detection, and Orphaned Object Analysis

.DESCRIPTION
    Enterprise-grade PowerShell automation tool that scans SCCM Collections across multiple sites
    to identify candidates for cleanup based on comprehensive criteria:

    COLLECTION ANALYSIS:
    - Collections not referenced by other collections
    - Collections not used as limiting collections
    - No assigned Power Configurations or Client Settings
    - Collections older than the -days threshold (default 365 for non-staging, 90 for staging)
    - Collections with retired/cancelled EPP (Endpoint Protection Package) packages
    - Active Directory group validation for user collections with whenChanged date tracking
    - Cross-validation against task sequence references

    RETIRED FOLDER DETECTION:
    - Automatically scans for objects in folders containing "Retired" in the path
    - Identifies User Collections, Device Collections, and Task Sequences in retired folders
    - Cross-validates retired objects against EPP package status
    - Flags objects as "Review Required" (collections) or "Discovered" (task sequences) if they reference
      active or unverifiable packages
    - Marks task sequences as "Archived" only when all referenced packages are confirmed retired/cancelled
    - Tracks referenced packages and collections for each retired task sequence

    STAGING ORPHANED OBJECT DETECTION:
    - Scans staging folders for collections with orphaned applications/baselines/packages
    - Identifies device collections with no active references
    - Identifies user collections with no active references
    - Detects task sequences referencing orphaned packages
    - Validates all associated objects for orphaned status
    - Excludes objects already processed in retired folders

    TASK SEQUENCE EPP CLEANUP:
    - Identifies task sequences referencing retired or cancelled EPP packages
    - Tracks referenced packages and associated collections
    - Excludes task sequences already in retired folders from duplicate processing
    - Cross-references with collection deployments
    - Validates EPP status with retry logic (up to 3 attempts)

    ADVANCED FEATURES:
    - Multi-site processing with parallel database connections
    - Buffered logging with automatic rotation and 30-day retention
    - Weighted progress tracking with time estimation
    - Real-time SQL query progress via InfoMessage event handlers
    - Comprehensive error tracking (Fatal, Recoverable, Warnings)
    - Resource management with automatic cleanup
    - SQL query performance monitoring with duration tracking
    - Parameterized database operations for security (SQL injection prevention)
    - EPP package status matching with intelligent retry logic and connection recovery
    - AD group change date validation for user collections (optional - graceful degradation)
    - Enhanced datetime handling for SQL Server compatibility (1753-9999 range validation)
    - Automatic log file cleanup (30-day retention)
    - Query export to .sql files for SSMS validation and troubleshooting

    OUTPUT:
    - Detailed logs with component-level tracking
    - Database records in monitoring tables:
      * tbl_SCCMCollections - Collection metadata with EPP status
      * tbl_SCCMApplications - Application/package details with EPP validation
      * tbl_SCCMTasksequences - Task sequence references
      * tbl_SCCMCollectionStatus - Processing status and recommended actions
      * tbl_SCCMTaskSequencesEPPCleanup - EPP cleanup candidates with match type tracking
      * tbl_SCCMsites - Site registration and timestamp tracking
    - Exported SQL queries for validation in SSMS (logs\CollectionDetectionQuery_*.sql)
    - Log analysis JSON export on error (logs\LogAnalysis_*.json)
    - Daily rotating logs with environment prefix (logs\SCCMCollectionScan_ENV_DATE.log)

.PARAMETER environment
    Target environment: dev, uat, or prod
    Determines which SCCM sites and databases to connect to based on XML configuration

    Required: Yes
    Type: String
    Validation: Must be one of: dev, uat, prod

.PARAMETER days
    Age threshold in days for collections with no member or AD group change (1-365)
    - Non-staging collections: this value (default 365)
    - Staging collections: 90 days, or this value when it is lower

    The default of 365 gives the original 365/90 behaviour.

.PARAMETER KeepHistory
    By default each scan replaces a site's previous results in the monitoring database, so
    objects that are no longer candidates do not linger. The old rows are removed only after
    the site's SCCM queries have succeeded. Specify -KeepHistory to skip that refresh and
    only add rows that are not already present.

    Required: No
    Type: Switch

    Required: Yes
    Type: Integer
    Validation: Must be between 1 and 365

.EXAMPLE
    .\SCCMCollectionScan.ps1 -environment uat -days 365

    Scans all UAT SCCM sites, identifies collections unchanged for 365+ days,
    validates against EPP package status, scans retired folders, and detects orphaned objects.

.EXAMPLE
    .\SCCMCollectionScan.ps1 -environment prod -days 365 -Verbose

    Runs in production with verbose logging enabled, showing detailed debug information
    including AD group lookups, EPP matching logic, SQL query progress, and real-time
    processing updates.

.INPUTS
    XML Configuration File (Library\Resources.xml):
    - SCCM site server names and database connections
    - SPDW (EPP) database connection details
    - Monitoring database connection details
    - Environment-specific settings
    - Adaptiva database configuration

.OUTPUTS
    Log Files (logs\):
    - SCCMCollectionScan_<ENV>_<DATE>.log - Daily rotating log with component tracking
    - LogAnalysis_<TIMESTAMP>.json - Error analysis export (if errors occur)
    - Automatic cleanup of logs older than 30 days

    SQL Query Files (logs\):
    - CollectionDetectionQuery_<TIMESTAMP>.sql - Exported query for SSMS validation

    Database Records (Monitoring Database):
    - Collection metadata with EPP status and AD group validation
    - Task sequence cleanup candidates with EPP match type
    - Retired folder object inventory with status classification
    - Orphaned object detection results from staging folders
    - Processing status and recommended actions
    - Site registration with timestamp tracking

.NOTES
    Created with:    Visual Studio Code
    Created on:     6/13/2022 10:51 AM
    Created by:     Mauro Cazabonnet
    Filename:       SCCMCollectionScan.ps1

    Major Refactor: 2/25/2026
    Version:        2.0

    ARCHITECTURE:
    - Object-oriented design using PowerShell classes
    - Modular function-based approach with single responsibility
    - Resource management with automatic cleanup
    - Comprehensive error handling and recovery
    - Event-driven SQL progress tracking

    DEPENDENCIES:
    - PowerShell 5.1 or higher
    - SQL Server connectivity (Windows Auth or SQL Auth)
    - Active Directory module (optional - graceful degradation if unavailable)
    - Library\Functions.ps1 (AD helper functions - dot-sourced at script level)
    - Library\Resources.xml (environment configuration)
    - System.DirectoryServices assembly (loaded automatically)

    SECURITY CONSIDERATIONS:
    - Uses parameterized SQL queries to prevent injection attacks
    - Supports both Windows and SQL authentication
    - SQL credentials for specific servers are hardcoded (SECURITY REVIEW REQUIRED)
    - Log files may contain sensitive collection/package names
    - Connection strings include retry logic for resilience
    - All database operations use parameterized commands

    PERFORMANCE:
    - Buffered logging (100 entries before flush)
    - Connection pooling and reuse
    - Optimized SQL queries with CTEs and indexes
    - Progress tracking with estimated completion time
    - Real-time SQL progress via InfoMessage events
    - Hashtable-based EPP lookups for O(1) performance
    - Batch processing with progress updates every 10 items
    - Query timeout handling with async execution monitoring

    ERROR HANDLING:
    - Three-tier error classification (Fatal, Recoverable, Warnings)
    - Automatic retry logic for EPP database connections (3 attempts, 15s delay)
    - Graceful degradation when AD module unavailable
    - Connection health checks before query execution
    - Comprehensive error logging with stack traces
    - Resource cleanup in finally blocks

    EXIT CODES:
    0  - Success (no fatal errors)
    1  - Fatal errors occurred
    2  - Recoverable errors occurred
    99 - Critical exception in main execution

    RECENT ENHANCEMENTS (v2.0):
    - Added staging orphaned object detection
    - Enhanced EPP validation with connection recovery
    - Real-time SQL progress tracking via InfoMessage events
    - Improved datetime handling for SQL Server compatibility
    - Enhanced AD group validation with error handling
    - Automatic log file cleanup (30-day retention)
    - Query export functionality for troubleshooting
    - Retired folder detection with EPP cross-validation
    - Exclusion logic to prevent duplicate processing
    - Enhanced status classification (Archived vs Review Required)
#>

using module .\Library\Classes.psm1

[CmdletBinding()]
Param (
    [Parameter(Mandatory = $true, Position = 1)]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("dev", "prod", "uat")]
    [string]$environment = "uat",

    [Parameter(Mandatory = $true, Position = 2)]
    [ValidateScript({$_ -gt 0 -and $_ -le 365 -and $_ -ne $null})]
    [int]$days = 365,

    [Parameter(Mandatory = $false)]
    [switch]$KeepHistory
)

#region Logging Framework
#endregion

#region Database Connection Manager
class DatabaseConnection {
    [string]$ServerName
    [string]$DatabaseName
    [string]$ConnectionString
    [System.Data.SqlClient.SqlConnection]$Connection
    [System.Data.SqlClient.SqlCommand]$Command
    [bool]$IsConnected = $false
    [string]$ConnectionType
    [bool]$UseSqlAuth = $false

    DatabaseConnection([string]$server, [string]$database, [string]$type) {
        $this.ServerName = $server
        $this.DatabaseName = $database
        $this.ConnectionType = $type

        # Servers needing SQL authentication come from SCCM_SQL_AUTH_SERVERS (regex); all others use Windows auth
        $sqlAuthPattern = [Environment]::GetEnvironmentVariable('SCCM_SQL_AUTH_SERVERS')
        if ($sqlAuthPattern -and $server -match $sqlAuthPattern) {
            $this.UseSqlAuth = $true

            # SQL Authentication connection string
            $this.ConnectionString = "Server=$server; Database=$database; " +
                                    "User Id=sa; " +
                                    "Password=$([Environment]::GetEnvironmentVariable('SCCM_SQL_SA_PASSWORD')); " +
                                    "Connection Timeout=60; " +
                                    "ConnectRetryCount=3; " +
                                    "ConnectRetryInterval=10; " +
                                    "MultipleActiveResultSets=True; " +
                                    "Encrypt=False; " +
                                    "TrustServerCertificate=True;"
        }
        else {
            # Windows Authentication (Integrated Security) for all other servers
            $this.ConnectionString = "Server=$server; Database=$database; Integrated Security=True; " +
                                    "Connection Timeout=60; " +
                                    "ConnectRetryCount=3; " +
                                    "ConnectRetryInterval=10; " +
                                    "MultipleActiveResultSets=True;"
        }
    }

    [bool] Connect([Logger]$logger, [ErrorTracker]$errorTracker) {
        try {
            $authType = if ($this.UseSqlAuth) { "SQL Authentication" } else { "Windows Authentication" }
            $logger.Info("Attempting to connect to $($this.ConnectionType) database: $($this.ServerName)\$($this.DatabaseName) using $authType", "DatabaseConnection")

            $this.Connection = New-Object System.Data.SqlClient.SqlConnection
            $this.Connection.ConnectionString = $this.ConnectionString

            $this.Command = New-Object System.Data.SqlClient.SqlCommand
            $this.Command.CommandTimeout = 0
            $this.Command.Connection = $this.Connection

            $this.Connection.Open()

            # Test connection
            $testCmd = $this.Connection.CreateCommand()
            $testCmd.CommandText = "SELECT 1"
            try {
                [void]$testCmd.ExecuteScalar()
                $this.IsConnected = $true
                $logger.Info("Successfully connected to $($this.ConnectionType) database using $authType", "DatabaseConnection")
                return $true
            }
            finally {
                $testCmd.Dispose()
            }
        }
        catch {
            $errorTracker.RecordError("DatabaseConnection",
                "Failed to connect to $($this.ConnectionType) database: $($this.ServerName)\$($this.DatabaseName)",
                $true,
                $_.Exception)
            $logger.Fatal("Database connection failed: $_", "DatabaseConnection")
            return $false
        }
    }

    [void] Disconnect([Logger]$logger) {
        try {
            if ($this.Command) {
                $this.Command.Dispose()
                $this.Command = $null
            }

            if ($this.Connection -and $this.Connection.State -eq 'Open') {
                $this.Connection.Close()
                $this.Connection.Dispose()
                $this.Connection = $null
                $this.IsConnected = $false
                $logger.Info("Disconnected from $($this.ConnectionType) database", "DatabaseConnection")
            }
        }
        catch {
            $logger.Error("Error during disconnect: $_", "DatabaseConnection")
        }
    }
}
#endregion

#region Resource Manager
class ResourceManager {
    [System.Collections.ArrayList]$Resources
    [Logger]$Logger

    ResourceManager([Logger]$logger) {
        $this.Resources = New-Object System.Collections.ArrayList
        $this.Logger = $logger
    }

    [void] Register([object]$resource, [string]$name) {
        $resourceInfo = @{
            Name = $name
            Resource = $resource
            RegisteredAt = Get-Date
        }
        [void]$this.Resources.Add($resourceInfo)
        $this.Logger.Debug("Registered resource: $name", "ResourceManager")
    }

    [void] CleanupAll() {
        $this.Logger.Info("Starting cleanup of $($this.Resources.Count) resources", "ResourceManager")

        foreach ($resourceInfo in $this.Resources) {
            try {
                $resource = $resourceInfo.Resource
                $name = $resourceInfo.Name
        if ($resource -is [DatabaseConnection]) {
            $resource.Disconnect($this.Logger)
        }
        elseif ($resource -is [System.Data.SqlClient.SqlDataAdapter]) {
            $resource.Dispose()
        }
        elseif ($resource -is [System.IDisposable]) {
            $resource.Dispose()
        }

        $this.Logger.Debug("Cleaned up resource: $name", "ResourceManager")
    }
    catch {
        $this.Logger.Warning("Failed to cleanup resource $($resourceInfo.Name): $_", "ResourceManager")
    }
}

$this.Resources.Clear()
$this.Logger.Info("Resource cleanup completed", "ResourceManager")
    }
}
#endregion

#region Directory Management
function Initialize-Directories {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ScriptConfig]$Config,

        [Parameter(Mandatory = $true)]
        [Logger]$Logger,

        [Parameter(Mandatory = $true)]
        [ErrorTracker]$ErrorTracker
    )

    try {
        $Logger.Info("Initializing directory structure", "DirectoryInit")

        $directories = @($Config.LogDir, $Config.ReportDir)

        foreach ($dir in $directories) {
            if (-not (Test-Path $dir)) {
                $Logger.Info("Creating directory: $dir", "DirectoryInit")
                New-Item -Type Directory -Path $dir -ErrorAction Stop | Out-Null
            } else {
                $Logger.Debug("Directory already exists: $dir", "DirectoryInit")
            }
        }

        $Logger.Info("Directory initialization completed successfully", "DirectoryInit")
        return $true
    }
    catch {
        $ErrorTracker.RecordError("DirectoryInit",
            "Failed to initialize directories",
            $true,
            $_.Exception)
        $Logger.Fatal("Directory initialization failed: $_", "DirectoryInit")
        return $false
    }
}
#endregion

#region Configuration Loading
function Get-ConfigurationData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ScriptConfig]$Config,

        [Parameter(Mandatory = $true)]
        [Logger]$Logger,

        [Parameter(Mandatory = $true)]
        [ErrorTracker]$ErrorTracker
    )

    try {
        $Logger.Info("Loading configuration from: $($Config.XmlPath)", "ConfigLoad")

        if (-not (Test-Path $Config.XmlPath)) {
            throw "Configuration file not found: $($Config.XmlPath)"
        }

        [xml]$xmlContent = Get-Content $Config.XmlPath -ErrorAction Stop

        $Logger.Info("Configuration loaded successfully", "ConfigLoad")

        return @{
            Success = $true
            XmlData = $xmlContent
        }
    }
    catch {
        $ErrorTracker.RecordError("ConfigLoad",
            "Failed to load configuration file",
            $true,
            $_.Exception)
        $Logger.Fatal("Configuration loading failed: $_", "ConfigLoad")

        return @{
            Success = $false
            XmlData = $null
        }
    }
}
#endregion

#region Environment Configuration Parser
function Get-EnvironmentConfiguration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [xml]$XmlData,

        [Parameter(Mandatory = $true)]
        [string]$Environment,

        [Parameter(Mandatory = $true)]
        [Logger]$Logger,

        [Parameter(Mandatory = $true)]
        [ErrorTracker]$ErrorTracker
    )

    try {
        $Logger.Info("Parsing environment configuration for: $Environment", "EnvConfig")

        $envLower = $Environment.ToLower()

        if ($envLower -notin @("prod", "uat", "dev")) {
            throw "Invalid environment: $Environment. Must be prod, uat, or dev."
        }

        # Environment is flagged by an attribute on each site node (e.g. uat="", prod="", prod2="")
        $sites = @($XmlData.SelectNodes("//*[@*]") | Where-Object {
            $_.LocalName -like 'site*' -and
            ($_.Attributes | Where-Object { $_.LocalName -match "^$envLower\d*$" })
        })
        $sqlspdw = $XmlData.SelectSingleNode("//*[@*[local-name()='spdwsql']]")
        $sqlmonitoring = $XmlData.SelectSingleNode("//*[@*[local-name()='monsql']]")

        if (-not $sites -or $sites.Count -eq 0) {
            throw "No sites found for environment: $Environment"
        }

        $config = @{
            Sites = $sites
            SPDWSqlServer = $sqlspdw.sqlserver
            SPDWSQLDBName = $sqlspdw.sqldbname
            MonSqlServer = $sqlmonitoring.sqlserver
            MonSQLDBName = $sqlmonitoring.sqldbname
            Environment = $Environment.ToUpper()
        }

        $Logger.Info("Found $($sites.Count) site(s) for environment $Environment", "EnvConfig")
        $Logger.Debug("SPDW Server: $($config.SPDWSqlServer)", "EnvConfig")
        $Logger.Debug("Monitoring Server: $($config.MonSqlServer)", "EnvConfig")

        return @{
            Success = $true
            Configuration = $config
        }
    }
    catch {
        $ErrorTracker.RecordError("EnvConfig",
            "Failed to parse environment configuration",
            $true,
            $_.Exception)
        $Logger.Fatal("Environment configuration parsing failed: $_", "EnvConfig")

        return @{
            Success = $false
            Configuration = $null
        }
    }
}
#endregion

#region Library Import
function Import-ScriptLibraries {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ScriptPath,

        [Parameter(Mandatory = $true)]
        [Logger]$Logger,

        [Parameter(Mandatory = $true)]
        [ErrorTracker]$ErrorTracker
    )

    try {
        $Logger.Info("Importing script libraries", "LibraryImport")

        # Load System.DirectoryServices assembly
        try {
            $Logger.Info("Loading System.DirectoryServices assembly", "LibraryImport")
            Add-Type -AssemblyName System.DirectoryServices -ErrorAction Stop
            $Logger.Info("System.DirectoryServices assembly loaded successfully", "LibraryImport")
        }
        catch {
            $Logger.Warning("Failed to load System.DirectoryServices assembly: $_", "LibraryImport")
        }

        # REMOVED: Don't dot-source here - return the path instead
        $functionsPath = Join-Path $ScriptPath "Library\Functions.ps1"

        if (-not (Test-Path $functionsPath)) {
            throw "Functions library not found: $functionsPath"
        }

        $Logger.Info("Functions file located: $functionsPath", "LibraryImport")

        # Return the path so it can be dot-sourced at script level
        return @{
            Success = $true
            FunctionsPath = $functionsPath
        }
    }
    catch {
        $ErrorTracker.RecordError("LibraryImport",
            "Failed to import script libraries",
            $false,
            $_.Exception)
        $Logger.Error("Library import failed: $_", "LibraryImport")

        return @{
            Success = $false
            FunctionsPath = $null
        }
    }
}
#endregion
#region Database Connection Initialization
function Initialize-DatabaseConnections {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$EnvConfig,

        [Parameter(Mandatory = $true)]
        [Logger]$Logger,

        [Parameter(Mandatory = $true)]
        [ErrorTracker]$ErrorTracker,

        [Parameter(Mandatory = $true)]
        [ResourceManager]$ResourceManager
    )

    try {
        $Logger.Info("Initializing database connections", "DBInit")

        # Create monitoring database connection
        $monConnection = [DatabaseConnection]::new(
            $EnvConfig.MonSqlServer,
            $EnvConfig.MonSQLDBName,
            "Monitoring"
        )

        if (-not $monConnection.Connect($Logger, $ErrorTracker)) {
            throw "Failed to connect to Monitoring database"
        }

        $ResourceManager.Register($monConnection, "MonitoringDB")

        # Create SPDW database connection
        $spdwConnection = [DatabaseConnection]::new(
            $EnvConfig.SPDWSqlServer,
            $EnvConfig.SPDWSQLDBName,
            "SPDW"
        )

        if (-not $spdwConnection.Connect($Logger, $ErrorTracker)) {
            throw "Failed to connect to SPDW database"
        }

        $ResourceManager.Register($spdwConnection, "SPDWDB")

        $Logger.Info("All database connections initialized successfully", "DBInit")

        return @{
            Success = $true
            MonitoringConnection = $monConnection
            SPDWConnection = $spdwConnection
        }
    }
    catch {
        $ErrorTracker.RecordError("DBInit",
            "Failed to initialize database connections",
            $true,
            $_.Exception)
        $Logger.Fatal("Database initialization failed: $_", "DBInit")

        return @{
            Success = $false
            MonitoringConnection = $null
            SPDWConnection = $null
        }
    }
}
#endregion

#region SCCM Site Connection
function Connect-SCCMSite {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$SiteItem,

        [Parameter(Mandatory = $true)]
        [Logger]$Logger,

        [Parameter(Mandatory = $true)]
        [ErrorTracker]$ErrorTracker,

        [Parameter(Mandatory = $true)]
        [ResourceManager]$ResourceManager
    )

    try {
        $Logger.Info("Connecting to SCCM site: $($SiteItem.site)", "SCCMConnect")

        $sccmConnection = [DatabaseConnection]::new(
            $SiteItem.sqlserver,
            $SiteItem.sccmsqldbname,
            "SCCM-$($SiteItem.site)"
        )

        if (-not $sccmConnection.Connect($Logger, $ErrorTracker)) {
            throw "Failed to connect to SCCM database for site: $($SiteItem.site)"
        }

        $ResourceManager.Register($sccmConnection, "SCCMDB-$($SiteItem.site)")

        $Logger.Info("Successfully connected to SCCM site: $($SiteItem.site)", "SCCMConnect")

        return @{
            Success = $true
            Connection = $sccmConnection
        }
    }
    catch {
        $ErrorTracker.RecordError("SCCMConnect",
            "Failed to connect to SCCM site: $($SiteItem.site)",
            $false,
            $_.Exception)
        $Logger.Error("SCCM site connection failed: $_", "SCCMConnect")

        return @{
            Success = $false
            Connection = $null
        }
    }
}
#endregion

#region Staging Orphaned Collection Scanner
function Invoke-StagingOrphanedCollectionScan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [DatabaseConnection]$SCCMConnection,

        [Parameter(Mandatory = $true)]
        [string]$CollectionType, # 'Device' or 'User'

        [Parameter(Mandatory = $true)]
        [Logger]$Logger,

        [Parameter(Mandatory = $true)]
        [ErrorTracker]$ErrorTracker,

        [Parameter(Mandatory = $true)]
        [ResourceManager]$ResourceManager,

        [Parameter(Mandatory = $false)]
        [ProgressTracker]$Progress,

        [Parameter(Mandatory = $false)]
        [int]$SiteIndex = 1,

        [Parameter(Mandatory = $false)]
        [int]$TotalSites = 1
    )

    try {
        $Logger.Info("Starting staging orphaned $CollectionType collection scan", "StagingOrphanScan")

        # Build and execute query
        $objectType = "${CollectionType}Collection"
        $query = New-OrphanedObjectDetectionQuery -ObjectType $objectType -Logger $Logger

        if ($Progress) {
            $Progress.UpdateProgress("ProcessSites",
                "Site $SiteIndex/$TotalSites - Scanning staging $CollectionType collections",
                (($SiteIndex - 1) / $TotalSites * 100) + 5)
        }

        $dataset = New-Object System.Data.DataSet
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($query, $SCCMConnection.Connection)
        $adapter.SelectCommand.CommandTimeout = 300

        $ResourceManager.Register($adapter, "StagingOrphan${CollectionType}Adapter")

        [void]$adapter.Fill($dataset)

        $collections = $dataset.Tables[0]

        $Logger.Info("Found $($collections.Rows.Count) staging $CollectionType collections", "StagingOrphanScan")

        # Now check each collection for orphaned applications/baselines
        $orphanedCollections = New-Object System.Collections.ArrayList

        foreach ($collection in $collections.Rows) {
            # Get associated applications/baselines for this collection
            $associatedObjects = Get-CollectionAssociatedObjects `
                -SCCMConnection $SCCMConnection `
                -CollectionID $collection.CollectionSiteID `
                -Logger $Logger

            $hasOrphanedObjects = $false
            $orphanedObjectsList = @()

            # Check each associated object
            foreach ($obj in $associatedObjects) {
                $refCheck = $null

                switch ($obj.ObjectType) {
                    'Application' {
                        $refCheck = Get-ApplicationReferences `
                            -SCCMConnection $SCCMConnection `
                            -ApplicationName $obj.ObjectName `
                            -Logger $Logger
                    }
                    'Baseline' {
                        $refCheck = Get-BaselineReferences `
                            -SCCMConnection $SCCMConnection `
                            -BaselineName $obj.ObjectName `
                            -Logger $Logger
                    }
                    'Package' {
                        $refCheck = Get-PackageReferences `
                            -SCCMConnection $SCCMConnection `
                            -PackageID $obj.ObjectID `
                            -Logger $Logger
                    }
                }

                if ($refCheck -and $refCheck.IsOrphaned) {
                    $hasOrphanedObjects = $true
                    $orphanedObjectsList += $obj
                }
            }

            if ($hasOrphanedObjects) {
                [void]$orphanedCollections.Add(@{
                    Collection = $collection
                    OrphanedObjects = $orphanedObjectsList
                })
            }
        }

        $Logger.Info("Identified $($orphanedCollections.Count) staging $CollectionType collections with orphaned objects", "StagingOrphanScan")

        return @{
            Success = $true
            TotalCollections = $collections.Rows.Count
            OrphanedCollections = $orphanedCollections
            CollectionType = $CollectionType
        }
    }
    catch {
        $ErrorTracker.RecordError("StagingOrphanScan",
            "Failed to scan staging $CollectionType collections for orphans",
            $false,
            $_.Exception)
        $Logger.Error("Staging orphaned collection scan failed: $_", "StagingOrphanScan")

        return @{
            Success = $false
            TotalCollections = 0
            OrphanedCollections = @()
            CollectionType = $CollectionType
        }
    }
}
#endregion
#region Staging Orphaned Task Sequence Scanner
function Invoke-StagingOrphanedTaskSequenceScan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [DatabaseConnection]$SCCMConnection,

        [Parameter(Mandatory = $true)]
        [Logger]$Logger,

        [Parameter(Mandatory = $true)]
        [ErrorTracker]$ErrorTracker,

        [Parameter(Mandatory = $true)]
        [ResourceManager]$ResourceManager,

        [Parameter(Mandatory = $false)]
        [ProgressTracker]$Progress,

        [Parameter(Mandatory = $false)]
        [int]$SiteIndex = 1,

        [Parameter(Mandatory = $false)]
        [int]$TotalSites = 1
    )

    try {
        $Logger.Info("Starting staging orphaned task sequence scan", "StagingTSOrphanScan")

        # Build and execute query
        $query = New-OrphanedObjectDetectionQuery -ObjectType 'TaskSequence' -Logger $Logger

        if ($Progress) {
            $Progress.UpdateProgress("ProcessSites",
                "Site $SiteIndex/$TotalSites - Scanning staging task sequences",
                (($SiteIndex - 1) / $TotalSites * 100) + 7)
        }

        $dataset = New-Object System.Data.DataSet
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($query, $SCCMConnection.Connection)
        $adapter.SelectCommand.CommandTimeout = 300

        $ResourceManager.Register($adapter, "StagingOrphanTSAdapter")

        [void]$adapter.Fill($dataset)

        $taskSequences = $dataset.Tables[0]

        $Logger.Info("Found $($taskSequences.Rows.Count) staging task sequences", "StagingTSOrphanScan")

        # Check each task sequence for orphaned packages
        $orphanedTS = New-Object System.Collections.ArrayList

        foreach ($ts in $taskSequences.Rows) {
            # Get referenced packages for this task sequence
            $referencedPackages = Get-TaskSequenceReferencedPackages `
                -SCCMConnection $SCCMConnection `
                -TaskSequenceID $ts.PackageID `
                -Logger $Logger

            $hasOrphanedPackages = $false
            $orphanedPackagesList = @()

            # Check each referenced package
            foreach ($pkg in $referencedPackages) {
                $refCheck = Get-PackageReferences `
                    -SCCMConnection $SCCMConnection `
                    -PackageID $pkg.PackageID `
                    -Logger $Logger

                if ($refCheck.IsOrphaned) {
                    $hasOrphanedPackages = $true
                    $orphanedPackagesList += $pkg
                }
            }

            if ($hasOrphanedPackages) {
                [void]$orphanedTS.Add(@{
                    TaskSequence = $ts
                    OrphanedPackages = $orphanedPackagesList
                })
            }
        }

        $Logger.Info("Identified $($orphanedTS.Count) staging task sequences with orphaned packages", "StagingTSOrphanScan")

        return @{
            Success = $true
            TotalTaskSequences = $taskSequences.Rows.Count
            OrphanedTaskSequences = $orphanedTS
        }
    }
    catch {
        $ErrorTracker.RecordError("StagingTSOrphanScan",
            "Failed to scan staging task sequences for orphans",
            $false,
            $_.Exception)
        $Logger.Error("Staging task sequence orphan scan failed: $_", "StagingTSOrphanScan")

        return @{
            Success = $false
            TotalTaskSequences = 0
            OrphanedTaskSequences = @()
        }
    }
}
#endregion

#region Functions
function Add-OrphanedObjectToDatabase {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$OrphanedObject,

        [Parameter(Mandatory = $true)]
        [string]$SiteId,

        [Parameter(Mandatory = $true)]
        [string]$ObjectType, # 'Collection' or 'TaskSequence'

        [Parameter(Mandatory = $true)]
        [System.Data.SqlClient.SqlCommand]$MonitoringCommand,

        [Parameter(Mandatory = $true)]
        [Logger]$Logger,

        [Parameter(Mandatory = $true)]
        [ErrorTracker]$ErrorTracker,

        [Parameter(Mandatory = $true)]
        [datetime]$Timestamp
    )

    try {
        if ($ObjectType -eq 'Collection') {
            $collection = $OrphanedObject.Collection
            $orphanedObjects = $OrphanedObject.OrphanedObjects

            # Sanitize values
            $collectionName = ConvertTo-SafeSqlString -Value $collection.CollectionName
            $collectionPath = ConvertTo-SafeSqlString -Value $collection.FolderPath
            $collectionType = if ($collection.PSObject.Properties.Name -contains 'CollectionType') {
                $collection.CollectionType
            } else {
                'Unknown'
            }

            # Build detailed orphaned objects breakdown
            $orphanedDetails = @()
            $appCount = 0
            $pkgCount = 0
            $baselineCount = 0

            foreach ($obj in $orphanedObjects) {
                $orphanedDetails += "- $($obj.ObjectType): '$($obj.ObjectName)' (ID: $($obj.ObjectID)) - No active references found"

                switch ($obj.ObjectType) {
                    'Application' { $appCount++ }
                    'Package' { $pkgCount++ }
                    'Baseline' { $baselineCount++ }
                }
            }

            # Build comprehensive message
            $detailedMessage = @"
STAGING FOLDER COLLECTION WITH ORPHANED REFERENCES DETECTED

Collection Details:
- Name: $collectionName
- Type: $collectionType
- Path: $collectionPath
- Member Count: $($collection.MemberCount)
- Days Since Last Change: $($collection.DaysSinceLastChange)

Referenced Objects with No Active Usage:
$($orphanedDetails -join "`n")

Summary:
- Total Orphaned Objects: $($orphanedObjects.Count)
- Applications: $appCount
- Packages: $pkgCount
- Baselines: $baselineCount

Detection Criteria:
These objects have no:
X Active deployments to collections
X References from other task sequences
X Active advertisements
X Distribution point group assignments

Reason for Orphaned Status:
All referenced objects are unused and can be considered for cleanup.
"@

            $Logger.Debug("Inserting orphaned collection: $collectionName with $($orphanedObjects.Count) orphaned objects", "OrphanedObjectInsert")

            # Insert into collections table with orphaned flag
            $sqlInsert = @"
INSERT INTO [dbo].[tbl_SCCMCollections]
    ([siteid],[collectionid],[collectionsiteid],[name],[path],[lastmemberchangetime],[lastmemberchangedays],
    [adaptivaclientsettingsref],[adaptivapolicyref],[adaptivahealthcheckref],[membercount],[collectiontype],
    [deploymenttype],[baselinetotaldevicecnt],[compliancecnt],[baselinecompliance],[baselinecollectionrefcnt],
    [baselineconfigitemscount],[baselineconfigitemreference],[baselinecollectionreference],
    [LimitingCollection],[ReferencedCollection],[timestamp])
SELECT @SiteId, @CollectionID, @CollectionSiteID, @Name, @Path, @LastMemberChangeTime, @LastMemberChangeDays,
    0, 0, 0, @MemberCount, @CollectionType,
    'None', 'N/A', 'N/A', 'N/A', 'N/A',
    'N/A', 'N/A', 'N/A',
    'False', 'False', @Timestamp
WHERE NOT EXISTS (
    SELECT [collectionid] FROM [dbo].[tbl_SCCMCollections]
    WHERE [collectionid] = @CollectionID
)
"@

            $MonitoringCommand.Parameters.Clear()

            [void]$MonitoringCommand.Parameters.AddWithValue("@SiteId", $SiteId)
            [void]$MonitoringCommand.Parameters.AddWithValue("@CollectionID", $collection.CollectionID)
            [void]$MonitoringCommand.Parameters.AddWithValue("@CollectionSiteID", $collection.CollectionSiteID)
            [void]$MonitoringCommand.Parameters.AddWithValue("@Name", $collectionName)
            [void]$MonitoringCommand.Parameters.AddWithValue("@Path", $collectionPath)

            $dateParam = $MonitoringCommand.Parameters.Add("@LastMemberChangeTime", [System.Data.SqlDbType]::DateTime)
            $dateParam.Value = $collection.LastMemberChangeTime

            [void]$MonitoringCommand.Parameters.AddWithValue("@LastMemberChangeDays", $collection.DaysSinceLastChange)
            [void]$MonitoringCommand.Parameters.AddWithValue("@MemberCount", $collection.MemberCount)
            [void]$MonitoringCommand.Parameters.AddWithValue("@CollectionType", $collectionType)
            [void]$MonitoringCommand.Parameters.AddWithValue("@Timestamp", $Timestamp)

            $MonitoringCommand.CommandText = $sqlInsert
            [void]$MonitoringCommand.ExecuteNonQuery()

            # Insert enhanced status record
            Add-OrphanedObjectStatus `
                -ObjectId $collection.CollectionID `
                -ObjectType 'Collection' `
                -DetailedMessage $detailedMessage `
                -OrphanedCount $orphanedObjects.Count `
                -MonitoringCommand $MonitoringCommand `
                -Logger $Logger `
                -ErrorTracker $ErrorTracker `
                -Timestamp $Timestamp

            $Logger.Info("Successfully inserted orphaned collection: $collectionName ($($orphanedObjects.Count) orphaned objects)", "OrphanedObjectInsert")
            return $true
        }
        elseif ($ObjectType -eq 'TaskSequence') {
            $taskSequence = $OrphanedObject.TaskSequence
            $orphanedPackages = $OrphanedObject.OrphanedPackages

            # Sanitize values
            $tsName = ConvertTo-SafeSqlString -Value $taskSequence.TaskSequenceName
            $tsPath = if ($taskSequence.PSObject.Properties.Name -contains 'FolderPath') {
                ConvertTo-SafeSqlString -Value $taskSequence.FolderPath
            } else {
                '/'
            }
            $packageId = ConvertTo-SafeSqlString -Value $taskSequence.PackageID

            # Build detailed orphaned packages breakdown
            $orphanedPackageDetails = @()
            $packageTypeBreakdown = @{}

            foreach ($pkg in $orphanedPackages) {
                $orphanedPackageDetails += "- $($pkg.PackageType): '$($pkg.PackageName)' (ID: $($pkg.PackageID)) - No active references"

                if ($packageTypeBreakdown.ContainsKey($pkg.PackageType)) {
                    $packageTypeBreakdown[$pkg.PackageType]++
                } else {
                    $packageTypeBreakdown[$pkg.PackageType] = 1
                }
            }

            # Build package type summary
            $packageTypeSummary = ($packageTypeBreakdown.GetEnumerator() | ForEach-Object {
                "- $($_.Key): $($_.Value)"
            }) -join "`n"

            # Build comprehensive message
            $detailedMessage = @"
STAGING FOLDER TASK SEQUENCE WITH ORPHANED PACKAGE REFERENCES DETECTED

Task Sequence Details:
- Name: $tsName
- Package ID: $packageId
- Path: $tsPath
- Days Since Last Change: $($taskSequence.DaysSinceLastChange)

Referenced Packages with No Active Usage:
$($orphanedPackageDetails -join "`n")

Summary:
- Total Orphaned Packages: $($orphanedPackages.Count)

Package Type Breakdown:
$packageTypeSummary

Detection Criteria:
These packages are not referenced by:
X Other task sequences
X Active advertisements
X Distribution point groups

Reason for Orphaned Status:
All referenced packages are unused and the task sequence can be considered for cleanup.
"@

            $Logger.Debug("Inserting orphaned task sequence: $tsName with $($orphanedPackages.Count) orphaned packages", "OrphanedObjectInsert")

            # Insert into task sequences table
            $sqlInsert = @"
INSERT INTO [dbo].[tbl_SCCMtasksequences]
    ([collectionid],[collectionsiteid],[name],[apptsref],[timestamp])
SELECT NULL, @PackageID, @Name, 0, @Timestamp
WHERE NOT EXISTS (
    SELECT [collectionsiteid],[name] FROM [dbo].[tbl_SCCMtasksequences]
    WHERE [collectionsiteid] = @PackageID
    AND [name] = @Name
)
"@

            $MonitoringCommand.Parameters.Clear()

            [void]$MonitoringCommand.Parameters.AddWithValue("@PackageID", $packageId)
            [void]$MonitoringCommand.Parameters.AddWithValue("@Name", $tsName)

            $timestampParam = $MonitoringCommand.Parameters.Add("@Timestamp", [System.Data.SqlDbType]::DateTime)
            $timestampParam.Value = $Timestamp

            $MonitoringCommand.CommandText = $sqlInsert
            [void]$MonitoringCommand.ExecuteNonQuery()

            # Insert enhanced status record
            Add-OrphanedObjectStatus `
                -ObjectId $packageId `
                -ObjectType 'TaskSequence' `
                -DetailedMessage $detailedMessage `
                -OrphanedCount $orphanedPackages.Count `
                -MonitoringCommand $MonitoringCommand `
                -Logger $Logger `
                -ErrorTracker $ErrorTracker `
                -Timestamp $Timestamp

            $Logger.Info("Successfully inserted orphaned task sequence: $tsName ($($orphanedPackages.Count) orphaned packages)", "OrphanedObjectInsert")
            return $true
        }
        return $false
    }
    catch {
        $ErrorTracker.RecordError("OrphanedObjectInsert",
            "Failed to insert orphaned object",
            $false,
            $_.Exception)
        $Logger.Error("Orphaned object insert failed: $_", "OrphanedObjectInsert")
        return $false
    }
}
function Add-OrphanedObjectStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ObjectId,

        [Parameter(Mandatory = $true)]
        [string]$ObjectType,

        [Parameter(Mandatory = $true)]
        [string]$OrphanedObjectsList,

        [Parameter(Mandatory = $true)]
        [System.Data.SqlClient.SqlCommand]$MonitoringCommand,

        [Parameter(Mandatory = $true)]
        [Logger]$Logger,

        [Parameter(Mandatory = $true)]
        [ErrorTracker]$ErrorTracker,

        [Parameter(Mandatory = $true)]
        [datetime]$Timestamp
    )

    try {
        $Logger.Debug("Inserting orphaned object status for: $ObjectId", "OrphanedStatusInsert")

        $message = "Staging folder object with orphaned references: $OrphanedObjectsList"
        $actions = "Review for cleanup - contains orphaned objects"

        $sqlInsert = @"
INSERT INTO [dbo].[tbl_SCCMCollectionStatus]
    ([collectionid],[status],[message],[actions],[timestamp])
SELECT @ObjectID, @Status, @Message, @Actions, @Timestamp
WHERE NOT EXISTS (
    SELECT [collectionid] FROM [dbo].[tbl_SCCMCollectionStatus]
    WHERE [collectionid] = @ObjectID
)
"@

        $MonitoringCommand.Parameters.Clear()

        [void]$MonitoringCommand.Parameters.AddWithValue("@ObjectID", $ObjectId)
        [void]$MonitoringCommand.Parameters.AddWithValue("@Status", "Orphaned")
        [void]$MonitoringCommand.Parameters.AddWithValue("@Message", $message)
        [void]$MonitoringCommand.Parameters.AddWithValue("@Actions", $actions)

        $timestampParam = $MonitoringCommand.Parameters.Add("@Timestamp", [System.Data.SqlDbType]::DateTime)
        $timestampParam.Value = $Timestamp

        $MonitoringCommand.CommandText = $sqlInsert
        [void]$MonitoringCommand.ExecuteNonQuery()

        $Logger.Debug("Successfully inserted orphaned object status", "OrphanedStatusInsert")
        return $true
    }
    catch {
        $ErrorTracker.RecordError("OrphanedStatusInsert",
            "Failed to insert orphaned object status for: $ObjectId",
            $false,
            $_.Exception)
        $Logger.Error("Orphaned object status insert failed: $_", "OrphanedStatusInsert")
        return $false
    }
}
function Add-RetiredCollectionToDatabase {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$RetiredItem,

        [Parameter(Mandatory = $true)]
        [string]$SiteId,

        [Parameter(Mandatory = $true)]
        [System.Data.SqlClient.SqlCommand]$MonitoringCommand,

        [Parameter(Mandatory = $true)]
        [Logger]$Logger,

        [Parameter(Mandatory = $true)]
        [ErrorTracker]$ErrorTracker,

        [Parameter(Mandatory = $true)]
        [datetime]$Timestamp,

        # NEW: EPP Data for validation
        [Parameter(Mandatory = $false)]
        [System.Data.DataTable]$EPPData = $null,

        # NEW: Task Sequence data to check references
        [Parameter(Mandatory = $false)]
        [System.Data.DataTable]$TaskSequenceData = $null
    )

    try {
        # Sanitize values
        $collectionName = ConvertTo-SafeSqlString -Value $RetiredItem.Name
        $collectionPath = ConvertTo-SafeSqlString -Value $RetiredItem.FolderPath
        $collectionType = $RetiredItem.ObjectType

        # NEW: Validate EPP status if collection references task sequences
        $status = "Archived"
        $message = "Found in retired folder during scan"
        $hasActivePackages = $false

        if ($null -ne $EPPData -and $null -ne $TaskSequenceData -and $TaskSequenceData.Rows.Count -gt 0) {
            $Logger.Debug("Checking EPP status for retired collection: $collectionName", "RetiredCollectionInsert")

            # Find task sequences that reference this collection
            $referencingTS = $TaskSequenceData.Rows | Where-Object {
                $_.CollectionID -eq $RetiredItem.CollectionID -or
                $_.CollectionName -eq $RetiredItem.Name
            }

            if ($referencingTS) {
                $Logger.Debug("Found $(@($referencingTS).Count) task sequence(s) referencing collection: $collectionName", "RetiredCollectionInsert")

                # Check EPP status of referenced packages
                foreach ($ts in $referencingTS) {
                    $refPackageName = $ts.'Referenced Package Name'

                    if (-not [string]::IsNullOrWhiteSpace($refPackageName)) {
                        $eppStatus = Get-EPPPackageStatus `
                            -EPPData $EPPData `
                            -PackageName $refPackageName `
                            -AppComments "" `
                            -Logger $Logger

                        # If any package is active (not Retired/Cancelled), flag it
                        if ($eppStatus.Status -notin @("Retired", "Cancel Package", "Not Found", "Null", "Error")) {
                            $hasActivePackages = $true
                            $Logger.Warning("Collection $collectionName references active package: $refPackageName (Status: $($eppStatus.Status))", "RetiredCollectionInsert")
                            break
                        }
                    }
                }
            }
        }

        # NEW: Adjust status based on EPP validation
        if ($hasActivePackages) {
            $status = "Review Required"
            $message = "In retired folder but references task sequences with active packages"
            $Logger.Warning("Collection $collectionName marked as 'Review Required' due to active package references", "RetiredCollectionInsert")
        }

        # Validate datetime
        $lastMemberChangeTime = $null
        $minSqlDate = [DateTime]::Parse("1753-01-01")
        $maxSqlDate = [DateTime]::Parse("9999-12-31")

        try {
            $testDate = [DateTime]::Parse($RetiredItem.LastMemberChangeTime.ToString())

            if ($testDate -ge $minSqlDate -and $testDate -le $maxSqlDate) {
                $lastMemberChangeTime = $testDate
            } else {
                $lastMemberChangeTime = Get-Date
            }
        }
        catch {
            $lastMemberChangeTime = Get-Date
        }

        $daysSinceLastChange = [int]$RetiredItem.DaysSinceLastChange

        $Logger.Debug("Inserting retired collection: $collectionName with status: $status", "RetiredCollectionInsert")

        # Use existing tbl_SCCMcollections schema
        $sqlInsert = @"
INSERT INTO [dbo].[tbl_SCCMcollections]
    ([siteid],[collectionid],[collectionsiteid],[name],[path],[lastmemberchangetime],[lastmemberchangedays],
    [adaptivaclientsettingsref],[adaptivapolicyref],[adaptivahealthcheckref],[membercount],[collectiontype],
    [deploymenttype],[baselinetotaldevicecnt],[compliancecnt],[baselinecompliance],[baselinecollectionrefcnt],
    [baselineconfigitemscount],[baselineconfigitemreference],[baselinecollectionreference],
    [LimitingCollection],[ReferencedCollection],[timestamp])
SELECT @SiteId, @CollectionID, @CollectionSiteID, @Name, @Path, @LastMemberChangeTime, @LastMemberChangeDays,
    0, 0, 0, @MemberCount, @CollectionType,
    'None', 'N/A', 'N/A', 'N/A', 'N/A',
    'N/A', 'N/A', 'N/A',
    @LimitingCollection, 'False', @Timestamp
WHERE NOT EXISTS (
    SELECT [collectionid] FROM [dbo].[tbl_SCCMcollections]
    WHERE [collectionid] = @CollectionID
)
"@

        $MonitoringCommand.Parameters.Clear()

        [void]$MonitoringCommand.Parameters.AddWithValue("@SiteId", $SiteId)
        [void]$MonitoringCommand.Parameters.AddWithValue("@CollectionID", $RetiredItem.CollectionID)
        [void]$MonitoringCommand.Parameters.AddWithValue("@CollectionSiteID", $RetiredItem.CollectionSiteID)
        [void]$MonitoringCommand.Parameters.AddWithValue("@Name", $collectionName)
        [void]$MonitoringCommand.Parameters.AddWithValue("@Path", $collectionPath)

        $dateParam = $MonitoringCommand.Parameters.Add("@LastMemberChangeTime", [System.Data.SqlDbType]::DateTime)
        $dateParam.Value = $lastMemberChangeTime

        [void]$MonitoringCommand.Parameters.AddWithValue("@LastMemberChangeDays", $daysSinceLastChange)
        [void]$MonitoringCommand.Parameters.AddWithValue("@MemberCount", $RetiredItem.MemberCount)
        [void]$MonitoringCommand.Parameters.AddWithValue("@CollectionType", $collectionType)
        $limitingValue = if ($RetiredItem.Table.Columns.Contains("LimitingCollectionName") -and $RetiredItem.LimitingCollectionName -isnot [System.DBNull] -and -not [string]::IsNullOrEmpty([string]$RetiredItem.LimitingCollectionName)) { [string]$RetiredItem.LimitingCollectionName } else { [System.DBNull]::Value }
        [void]$MonitoringCommand.Parameters.AddWithValue("@LimitingCollection", $limitingValue)
        [void]$MonitoringCommand.Parameters.AddWithValue("@Timestamp", $Timestamp)

        $MonitoringCommand.CommandText = $sqlInsert
        [void]$MonitoringCommand.ExecuteNonQuery()

        # Insert status with validated status
        Add-RetiredCollectionStatus `
            -CollectionId $RetiredItem.CollectionID `
            -MonitoringCommand $MonitoringCommand `
            -Logger $Logger `
            -ErrorTracker $ErrorTracker `
            -Timestamp $Timestamp `
            -Status $status `
            -Message $message

        $Logger.Info("Successfully inserted retired collection: $collectionName with status: $status", "RetiredCollectionInsert")
        return $true
    }
    catch {
        $ErrorTracker.RecordError("RetiredCollectionInsert",
            "Failed to insert retired collection: $($RetiredItem.Name)",
            $false,
            $_.Exception)
        $Logger.Error("Retired collection insert failed: $_", "RetiredCollectionInsert")
        return $false
    }
}
function Add-RetiredTaskSequenceToDatabase {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$RetiredItem,

        [Parameter(Mandatory = $true)]
        [string]$SiteId,

        [Parameter(Mandatory = $true)]
        [System.Data.SqlClient.SqlCommand]$MonitoringCommand,

        [Parameter(Mandatory = $true)]
        [Logger]$Logger,

        [Parameter(Mandatory = $true)]
        [ErrorTracker]$ErrorTracker,

        [Parameter(Mandatory = $true)]
        [datetime]$Timestamp,

        [Parameter(Mandatory = $false)]
        [System.Data.DataTable]$EPPData = $null,

        # Numeric IDs of every collection this task sequence is deployed to
        [Parameter(Mandatory = $false)]
        [string[]]$CollectionIds = @()
    )

    try {
        # Sanitize basic values
        $tsName = ConvertTo-SafeSqlString -Value $RetiredItem.Name
        $tsPath = ConvertTo-SafeSqlString -Value $RetiredItem.FolderPath
        $packageId = ConvertTo-SafeSqlString -Value $RetiredItem.PackageID

        # Sanitize referenced package and collection values
        $refPackageId = if ($RetiredItem.PSObject.Properties.Name -contains 'ReferencePackageID' -and
            -not [string]::IsNullOrWhiteSpace($RetiredItem.ReferencePackageID)) {
            ConvertTo-SafeSqlString -Value $RetiredItem.ReferencePackageID
        } else { "N/A" }

        $refPackageName = if ($RetiredItem.PSObject.Properties.Name -contains 'ReferencedPackageName' -and
            -not [string]::IsNullOrWhiteSpace($RetiredItem.ReferencedPackageName)) {
            ConvertTo-SafeSqlString -Value $RetiredItem.ReferencedPackageName
        } else { "N/A" }

        $collectionName = if ($RetiredItem.PSObject.Properties.Name -contains 'CollectionName' -and
            -not [string]::IsNullOrWhiteSpace($RetiredItem.CollectionName)) {
            ConvertTo-SafeSqlString -Value $RetiredItem.CollectionName
        } else { "" }

        $collectionId = if ($RetiredItem.PSObject.Properties.Name -contains 'CollectionID' -and
            $null -ne $RetiredItem.CollectionID -and
            -not [string]::IsNullOrWhiteSpace($RetiredItem.CollectionID.ToString())) {
            ConvertTo-SafeSqlString -Value $RetiredItem.CollectionID
        } else { "" }

        $collectionType = if ($RetiredItem.PSObject.Properties.Name -contains 'CollectionType' -and
            -not [string]::IsNullOrWhiteSpace($RetiredItem.CollectionType)) {
            ConvertTo-SafeSqlString -Value $RetiredItem.CollectionType
        } else { "" }

        $collectionPath = if ($RetiredItem.PSObject.Properties.Name -contains 'CollectionObjectPath' -and
            -not [string]::IsNullOrWhiteSpace($RetiredItem.CollectionObjectPath)) {
            ConvertTo-SafeSqlString -Value $RetiredItem.CollectionObjectPath
        } else { "" }

        # Use ONLY valid CHECK constraint values
        $tsStatus = "Archived"  # Default to Archived
        $collStatus = if ($collectionName -ne "") { "Archived" } else { "None" }
        $tsMessage = "Found in retired folder during scan"
        $collMessage = if ($collectionName -ne "") { "Initial discovery scan" } else { "None" }
        $eppStatusValue = "Archived"
        $requiresReview = $false  # ✅ NEW: Track if review is needed

        if ($null -ne $EPPData -and $refPackageName -ne "N/A") {
            $Logger.Debug("Checking EPP status for retired task sequence: $tsName (Package: $refPackageName)", "RetiredTaskSequenceInsert")

            $eppStatus = Get-EPPPackageStatus `
                -EPPData $EPPData `
                -PackageName $refPackageName `
                -AppComments "" `
                -Logger $Logger

            # Only a confirmed Retired/Cancel Package status may be archived.
            # Active packages and unverifiable statuses (Not Found, Null, Error) both need review.
            if ($eppStatus.Status -notin @("Retired", "Cancel Package")) {
                $isUnverified = $eppStatus.Status -in @("Not Found", "Null", "Error")
                $tsStatus = "Discovered"
                $collStatus = if ($collectionName -ne "") { "Discovered" } else { "None" }
                $requiresReview = $true

                $reviewReason = if ($isUnverified) {
                    "EPP status could not be verified for package: $refPackageName (EPP Status: $($eppStatus.Status))"
                } else {
                    "In retired folder but uses ACTIVE package: $refPackageName (EPP Status: $($eppStatus.Status))"
                }
                $tsMessage = "⚠️ REQUIRES REVIEW: $reviewReason"
                $collMessage = if ($collectionName -ne "") {
                    "⚠️ REQUIRES REVIEW: Associated with task sequence whose package is active or unverified"
                } else {
                    "None"
                }
                $eppStatusValue = $eppStatus.Status

                $Logger.Debug("Task sequence $tsName marked as 'Discovered' (REQUIRES REVIEW) - $reviewReason", "RetiredTaskSequenceInsert")
            } else {
                $eppStatusValue = $eppStatus.Status
                $Logger.Info("Task sequence $tsName marked as 'Archived' - referenced package retired/cancelled", "RetiredTaskSequenceInsert")
            }
        }
        elseif ($null -eq $EPPData -and $refPackageName -ne "N/A") {
            # No EPP data at all: cannot confirm anything, so never archive
            $tsStatus = "Discovered"
            $collStatus = if ($collectionName -ne "") { "Discovered" } else { "None" }
            $requiresReview = $true
            $eppStatusValue = "Error"
            $tsMessage = "⚠️ REQUIRES REVIEW: EPP data unavailable - package not verified: $refPackageName"
            $collMessage = if ($collectionName -ne "") { "⚠️ REQUIRES REVIEW: EPP data unavailable" } else { "None" }
        }

        $Logger.Debug("Inserting retired task sequence: $tsName with status: $tsStatus (RefPkg: $refPackageName, EPP: $eppStatusValue, Review: $requiresReview)", "RetiredTaskSequenceInsert")

        # Link the task sequence and its package row to every collection it is deployed to
        # (-1 only when it has no deployment, or the collection could not be resolved)
        $linkIds = @($CollectionIds | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ($linkIds.Count -eq 0) { $linkIds = @("-1") }

        # 1. Insert into tbl_SCCMtasksequences
        $sqlInsert = @"
INSERT INTO [dbo].[tbl_SCCMtasksequences]
    ([collectionid],[collectionsiteid],[name],[apptsref],[timestamp])
SELECT @LinkCollectionID, @PackageID, @Name, 0, @Timestamp
WHERE NOT EXISTS (
    SELECT [collectionsiteid],[name] FROM [dbo].[tbl_SCCMtasksequences]
    WHERE [collectionsiteid] = @PackageID
    AND [name] = @Name
    AND [collectionid] = @LinkCollectionID
)
"@

        foreach ($linkId in $linkIds) {
            $MonitoringCommand.Parameters.Clear()

            [void]$MonitoringCommand.Parameters.AddWithValue("@LinkCollectionID", $linkId)
            [void]$MonitoringCommand.Parameters.AddWithValue("@PackageID", $packageId)
            [void]$MonitoringCommand.Parameters.AddWithValue("@Name", $tsName)

            $timestampParam = $MonitoringCommand.Parameters.Add("@Timestamp", [System.Data.SqlDbType]::DateTime)
            $timestampParam.Value = $Timestamp

            $MonitoringCommand.CommandText = $sqlInsert
            [void]$MonitoringCommand.ExecuteNonQuery()
        }

        # 2. Insert into tbl_SCCMapplications with validated EPP status
        $sqlAppInsert = @"
INSERT INTO [dbo].[tbl_SCCMapplications]
    ([collectionid],[collectionsiteid],[name],[path],[packageid],[packagesizeKB],[packagetype],
    [App_Comments],[apppref],[eppstatus],[timestamp])
SELECT @LinkCollectionID, @PackageID, @Name, @Path, @PackageID, 'N/A', @PackageType,
    @AppComments, 0, @EPPStatus, @Timestamp
WHERE NOT EXISTS (
    SELECT [packageid] FROM [dbo].[tbl_SCCMapplications]
    WHERE [packageid] = @PackageID
        AND [path] = @Path
        AND [collectionid] = @LinkCollectionID
)
"@

        # Enhanced comments with review flag
        $appCommentsText = if ($requiresReview) {
            "⚠️ REQUIRES REVIEW - Retired Folder - EPP Status: $eppStatusValue (ACTIVE)"
        } else {
            "Retired Folder - EPP Status: $eppStatusValue"
        }

        foreach ($linkId in $linkIds) {
            $MonitoringCommand.Parameters.Clear()

            [void]$MonitoringCommand.Parameters.AddWithValue("@LinkCollectionID", $linkId)
            [void]$MonitoringCommand.Parameters.AddWithValue("@PackageID", $packageId)
            [void]$MonitoringCommand.Parameters.AddWithValue("@Name", $tsName)
            [void]$MonitoringCommand.Parameters.AddWithValue("@Path", $tsPath)
            [void]$MonitoringCommand.Parameters.AddWithValue("@PackageType", "Task Sequence Package")
            [void]$MonitoringCommand.Parameters.AddWithValue("@AppComments", $appCommentsText)
            [void]$MonitoringCommand.Parameters.AddWithValue("@EPPStatus", $eppStatusValue)

            $timestampParam = $MonitoringCommand.Parameters.Add("@Timestamp", [System.Data.SqlDbType]::DateTime)
            $timestampParam.Value = $Timestamp

            $MonitoringCommand.CommandText = $sqlAppInsert
            [void]$MonitoringCommand.ExecuteNonQuery()
        }

        # 3. Insert into tbl_SCCMTaskSequencesEPPCleanup with validated statuses
        $sqlEPPCleanup = @"
INSERT INTO [dbo].[tbl_SCCMTaskSequencesEPPCleanup]
    ([siteid], [Task Sequence Name], [Task Sequence Package ID], [Task Sequence Path],
    [Referenced Package Name], [Reference Package ID], [Collection Name], [Collection ID],
    [Collection Type], [Collection Path], [EPP Status], [Match Type], [Timestamp],
    [Task Sequence Message], [Collection Message], [Task Sequence Status], [Collection Status],
    [Actions], [Statustimestamp])
SELECT @SiteId, @TSName, @TSPackageID, @TSPath,
    @RefPackageName, @RefPackageID, @CollectionName, @CollectionID,
    @CollectionType, @CollectionPath, @EPPStatus, @MatchType, @Timestamp,
    @TSMessage, @CollectionMessage,
    @TSStatus, @CollectionStatus,
    @Actions, @StatusTimestamp
WHERE NOT EXISTS (
    SELECT 1 FROM [dbo].[tbl_SCCMTaskSequencesEPPCleanup]
    WHERE [siteid] = @SiteId
    AND [Task Sequence Package ID] = @TSPackageID
    AND [Task Sequence Path] = @TSPath
    AND [Referenced Package Name] = @RefPackageName
    AND ([Collection ID] = @CollectionID OR (@CollectionID = '' AND [Collection ID] = ''))
)
"@

        $MonitoringCommand.Parameters.Clear()

        [void]$MonitoringCommand.Parameters.AddWithValue("@SiteId", $SiteId)
        [void]$MonitoringCommand.Parameters.AddWithValue("@TSName", $tsName)
        [void]$MonitoringCommand.Parameters.AddWithValue("@TSPackageID", $packageId)
        [void]$MonitoringCommand.Parameters.AddWithValue("@TSPath", $tsPath)
        [void]$MonitoringCommand.Parameters.AddWithValue("@RefPackageName", $refPackageName)
        [void]$MonitoringCommand.Parameters.AddWithValue("@RefPackageID", $refPackageId)
        [void]$MonitoringCommand.Parameters.AddWithValue("@CollectionName", $collectionName)
        [void]$MonitoringCommand.Parameters.AddWithValue("@CollectionID", $collectionId)
        [void]$MonitoringCommand.Parameters.AddWithValue("@CollectionType", $collectionType)
        [void]$MonitoringCommand.Parameters.AddWithValue("@CollectionPath", $collectionPath)
        [void]$MonitoringCommand.Parameters.AddWithValue("@EPPStatus", $eppStatusValue)
        [void]$MonitoringCommand.Parameters.AddWithValue("@MatchType", "Retired Folder")

        $timestampParam = $MonitoringCommand.Parameters.Add("@Timestamp", [System.Data.SqlDbType]::DateTime)
        $timestampParam.Value = $Timestamp

        [void]$MonitoringCommand.Parameters.AddWithValue("@TSMessage", $tsMessage)
        [void]$MonitoringCommand.Parameters.AddWithValue("@CollectionMessage", $collMessage)

        # Use validated status values
        [void]$MonitoringCommand.Parameters.AddWithValue("@TSStatus", $tsStatus)  # Will be "Discovered" or "Archived"
        [void]$MonitoringCommand.Parameters.AddWithValue("@CollectionStatus", $collStatus)

        # Enhanced actions based on review flag
        $actionsText = if ($requiresReview) {
            "⚠️ URGENT REVIEW REQUIRED - Active package in retired folder"
        } elseif ($tsStatus -eq "Archived") {
            "Review for deletion - all packages retired"
        } else {
            "Review for cleanup"
        }
        [void]$MonitoringCommand.Parameters.AddWithValue("@Actions", $actionsText)

        $statusTimestampParam = $MonitoringCommand.Parameters.Add("@StatusTimestamp", [System.Data.SqlDbType]::DateTime)
        $statusTimestampParam.Value = $Timestamp

        $MonitoringCommand.CommandText = $sqlEPPCleanup
        [void]$MonitoringCommand.ExecuteNonQuery()

        $Logger.Info("Successfully inserted retired task sequence: $tsName (Status: $tsStatus, RefPkg: $refPackageName, EPP: $eppStatusValue, Review: $requiresReview)", "RetiredTaskSequenceInsert")
        return $true
    }
    catch {
        $ErrorTracker.RecordError("RetiredTaskSequenceInsert",
            "Failed to insert retired task sequence: $($RetiredItem.Name)",
            $false,
            $_.Exception)
        $Logger.Error("Retired task sequence insert failed: $_", "RetiredTaskSequenceInsert")
        return $false
    }
}
function Add-RetiredCollectionStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$CollectionId,

        [Parameter(Mandatory = $true)]
        [System.Data.SqlClient.SqlCommand]$MonitoringCommand,

        [Parameter(Mandatory = $true)]
        [Logger]$Logger,

        [Parameter(Mandatory = $true)]
        [ErrorTracker]$ErrorTracker,

        [Parameter(Mandatory = $true)]
        [datetime]$Timestamp,

        # NEW: Allow custom status (default to "Archived")
        [Parameter(Mandatory = $false)]
        [string]$Status = "Archived",

        # NEW: Allow custom message
        [Parameter(Mandatory = $false)]
        [string]$Message = "Found in retired folder during scan"
    )

    try {
        $Logger.Debug("Inserting retired collection status for: $CollectionId (Status: $Status)", "RetiredStatusInsert")

        # Determine actions based on status
        $actions = if ($Status -eq "Review Required") {
            "Review - Active package references detected"
        } else {
            "Review for deletion"
        }

        $sqlInsert = @"
INSERT INTO [dbo].[tbl_SCCMCollectionStatus]
    ([collectionid],[status],[message],[actions],[timestamp])
SELECT @CollectionID, @Status, @Message, @Actions, @Timestamp
WHERE NOT EXISTS (
    SELECT [collectionid] FROM [dbo].[tbl_SCCMCollectionStatus]
    WHERE [collectionid] = @CollectionID
)
"@

        $MonitoringCommand.Parameters.Clear()

        [void]$MonitoringCommand.Parameters.AddWithValue("@CollectionID", $CollectionId)
        [void]$MonitoringCommand.Parameters.AddWithValue("@Status", $Status)
        [void]$MonitoringCommand.Parameters.AddWithValue("@Message", $Message)
        [void]$MonitoringCommand.Parameters.AddWithValue("@Actions", $actions)

        $timestampParam = $MonitoringCommand.Parameters.Add("@Timestamp", [System.Data.SqlDbType]::DateTime)
        $timestampParam.Value = $Timestamp

        $MonitoringCommand.CommandText = $sqlInsert
        [void]$MonitoringCommand.ExecuteNonQuery()

        $Logger.Debug("Successfully inserted retired collection status with status: $Status", "RetiredStatusInsert")
        return $true
    }
    catch {
        $ErrorTracker.RecordError("RetiredStatusInsert",
            "Failed to insert retired collection status for: $CollectionId",
            $false,
            $_.Exception)
        $Logger.Error("Retired collection status insert failed: $_", "RetiredStatusInsert")
        return $false
    }
}
function Add-CollectionToDatabase {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$CollectionItem,

        [Parameter(Mandatory = $true)]
        [string]$SiteId,

        [Parameter(Mandatory = $true)]
        [System.Data.SqlClient.SqlCommand]$MonitoringCommand,

        [Parameter(Mandatory = $true)]
        [Logger]$Logger,

        [Parameter(Mandatory = $true)]
        [ErrorTracker]$ErrorTracker,

        [Parameter(Mandatory = $true)]
        [datetime]$Timestamp,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$WhenChangedString = $null,

        [Parameter(Mandatory = $false)]
        [int]$DaysSinceChange = 0
    )

    try {
        # Sanitize all string values
        $collectionName = ConvertTo-SafeSqlString -Value $CollectionItem.CollectionName
        $collectionPath = ConvertTo-SafeSqlString -Value $CollectionItem.CollectionPath

        # Convert boolean values
        $adaptivaclientref = [System.Convert]::ToBoolean($CollectionItem."Adaptiva Client Settings Collection Reference")
        $adaptivapolicyref = [System.Convert]::ToBoolean($CollectionItem."Adaptiva Policy Collection Reference")
        $adaptivahealthref = [System.Convert]::ToBoolean($CollectionItem."Adaptiva Health Check Collection Reference")

        # CRITICAL FIX: Enhanced datetime validation and handling
        $lastMemberChangeTime = $null
        $daysSinceLastChange = 0
        $minSqlDate = [DateTime]::Parse("1753-01-01")
        $maxSqlDate = [DateTime]::Parse("9999-12-31")

        # Try to use AD group date first (if provided)
        if (-not [string]::IsNullOrWhiteSpace($WhenChangedString)) {
            try {
                $parsedDate = [DateTime]::Parse($WhenChangedString)

                # Validate SQL Server datetime range
                if ($parsedDate -ge $minSqlDate -and $parsedDate -le $maxSqlDate) {
                    $lastMemberChangeTime = $parsedDate
                    $daysSinceLastChange = $DaysSinceChange
                    $Logger.Debug("Using AD group date: $($lastMemberChangeTime.ToString('yyyy-MM-dd HH:mm:ss')) ($daysSinceLastChange days)", "CollectionInsert")
                } else {
                    $Logger.Warning("AD group date out of SQL range: $WhenChangedString", "CollectionInsert")
                }
            }
            catch {
                $Logger.Warning("Invalid AD group datetime format: $WhenChangedString - Error: $_", "CollectionInsert")
            }
        }

        # Fallback to collection date with comprehensive validation
        if ($null -eq $lastMemberChangeTime) {
            try {
                $collectionDate = $CollectionItem."LastMemberChangeTime"
                $testDate = $null

                # Handle different date formats and types
                if ($collectionDate -is [DateTime]) {
                    $testDate = $collectionDate
                }
                elseif (-not [string]::IsNullOrWhiteSpace($collectionDate)) {
                    try {
                        $testDate = [DateTime]::Parse($collectionDate.ToString())
                    }
                    catch {
                        $Logger.Warning("Failed to parse collection date: $collectionDate", "CollectionInsert")
                        $testDate = $null
                    }
                }

    # If still no valid date, use current timestamp
    if ($null -eq $testDate) {
        $testDate = Get-Date
        $Logger.Warning("No valid LastMemberChangeTime - using current timestamp for: $collectionName", "CollectionInsert")
    }

    # Validate SQL Server datetime range
    if ($testDate -lt $minSqlDate) {
        $Logger.Warning("Date too old - using minimum SQL date for: $collectionName", "CollectionInsert")
        $testDate = $minSqlDate
    }
    elseif ($testDate -gt $maxSqlDate) {
        $Logger.Warning("Date too far in future - using current date for: $collectionName", "CollectionInsert")
        $testDate = Get-Date
    }

    $lastMemberChangeTime = $testDate

    # Calculate days since change safely
    try {
        $daysValue = $CollectionItem."Days Since Last Member Change"

        if ($null -ne $daysValue) {
            $daysSinceLastChange = [int]$daysValue

            # Validate days value
            if ($daysSinceLastChange -lt 0 -or $daysSinceLastChange -gt 100000) {
                $daysSinceLastChange = ((Get-Date) - $testDate).Days
            }
        }
        else {
            $daysSinceLastChange = ((Get-Date) - $testDate).Days
        }
    }
    catch {
        $daysSinceLastChange = ((Get-Date) - $testDate).Days
    }
}
catch {
    # Absolute fallback
    $lastMemberChangeTime = Get-Date
    $daysSinceLastChange = 0
    $Logger.Warning("Error processing date - using current timestamp for: $collectionName", "CollectionInsert")
}
        }

# Final validation
if ($null -eq $lastMemberChangeTime) {
    $lastMemberChangeTime = Get-Date
    $daysSinceLastChange = 0
    $Logger.Warning("Final fallback - using current timestamp for: $collectionName", "CollectionInsert")
}

$Logger.Debug("Final datetime: $($lastMemberChangeTime.ToString('yyyy-MM-dd HH:mm:ss')), Days: $daysSinceLastChange", "CollectionInsert")

# SOLUTION: Use parameterized query instead of string concatenation
$sqlInsert = @"
INSERT INTO [dbo].[tbl_SCCMCollections]
    ([siteid],[collectionid],[collectionsiteid],[name],[path],[lastmemberchangetime],[lastmemberchangedays],
    [adaptivaclientsettingsref],[adaptivapolicyref],[adaptivahealthcheckref],[membercount],[collectiontype],
    [deploymenttype],[baselinetotaldevicecnt],[compliancecnt],[baselinecompliance],[baselinecollectionrefcnt],
    [baselineconfigitemscount],[baselineconfigitemreference],[baselinecollectionreference],
    [LimitingCollection],[ReferencedCollection],[timestamp])
SELECT @SiteId, @CollectionID, @CollectionSiteID, @Name, @Path, @LastMemberChangeTime, @LastMemberChangeDays,
       @AdaptivaClientRef, @AdaptivaPolicyRef, @AdaptivaHealthRef, @MemberCount, @CollectionType,
       @DeploymentType, @BaselineTotalDeviceCnt, @ComplianceCnt, @BaselineCompliance, @BaselineCollectionRefCnt,
       @BaselineConfigItemsCount, @BaselineConfigItemReference, @BaselineCollectionReference,
       @LimitingCollection, @ReferencedCollection, @Timestamp
WHERE NOT EXISTS (
    SELECT [collectionid] FROM [dbo].[tbl_SCCMCollections]
    WHERE [collectionid] = @CollectionID
)
"@

        # Clear existing parameters
        $MonitoringCommand.Parameters.Clear()

        # Add parameters with proper types
        [void]$MonitoringCommand.Parameters.AddWithValue("@SiteId", $SiteId)
        [void]$MonitoringCommand.Parameters.AddWithValue("@CollectionID", $CollectionItem.CollectionID)
        [void]$MonitoringCommand.Parameters.AddWithValue("@CollectionSiteID", $CollectionItem.CollectionSiteID)
        [void]$MonitoringCommand.Parameters.AddWithValue("@Name", $collectionName)
        [void]$MonitoringCommand.Parameters.AddWithValue("@Path", $collectionPath)

        # KEY FIX: Add datetime as DateTime object, not string
        $dateParam = $MonitoringCommand.Parameters.Add("@LastMemberChangeTime", [System.Data.SqlDbType]::DateTime)
        $dateParam.Value = $lastMemberChangeTime

        [void]$MonitoringCommand.Parameters.AddWithValue("@LastMemberChangeDays", $daysSinceLastChange)
        [void]$MonitoringCommand.Parameters.AddWithValue("@AdaptivaClientRef", $adaptivaclientref)
        [void]$MonitoringCommand.Parameters.AddWithValue("@AdaptivaPolicyRef", $adaptivapolicyref)
        [void]$MonitoringCommand.Parameters.AddWithValue("@AdaptivaHealthRef", $adaptivahealthref)
        [void]$MonitoringCommand.Parameters.AddWithValue("@MemberCount", $CollectionItem.MemberCount)
        [void]$MonitoringCommand.Parameters.AddWithValue("@CollectionType", $CollectionItem.CollectionType)
        [void]$MonitoringCommand.Parameters.AddWithValue("@DeploymentType", $CollectionItem.DeploymentType)
        [void]$MonitoringCommand.Parameters.AddWithValue("@BaselineTotalDeviceCnt", $CollectionItem.baselinetotaldevicecnt)
        [void]$MonitoringCommand.Parameters.AddWithValue("@ComplianceCnt", $CollectionItem.compliancecnt)
        [void]$MonitoringCommand.Parameters.AddWithValue("@BaselineCompliance", $CollectionItem.baselinecompliance)
        [void]$MonitoringCommand.Parameters.AddWithValue("@BaselineCollectionRefCnt", $CollectionItem.baselinecollectionrefcnt)
        [void]$MonitoringCommand.Parameters.AddWithValue("@BaselineConfigItemsCount", $CollectionItem.baselineconfigitemscount)
        [void]$MonitoringCommand.Parameters.AddWithValue("@BaselineConfigItemReference", $CollectionItem.baselineconfigitemreference)
        [void]$MonitoringCommand.Parameters.AddWithValue("@BaselineCollectionReference", $CollectionItem.baselinecollectionreference)
        [void]$MonitoringCommand.Parameters.AddWithValue("@LimitingCollection", $CollectionItem.LimitingCollection)
        [void]$MonitoringCommand.Parameters.AddWithValue("@ReferencedCollection", $CollectionItem.ReferencedCollection)
        [void]$MonitoringCommand.Parameters.AddWithValue("@Timestamp", $Timestamp)

        $MonitoringCommand.CommandText = $sqlInsert
        [void]$MonitoringCommand.ExecuteNonQuery()

        $Logger.Info("Successfully inserted collection: $collectionName", "CollectionInsert")
        return $true
    }
    catch {
        $ErrorTracker.RecordError("CollectionInsert",
            "Failed to insert collection: $($CollectionItem.CollectionName)",
            $false,
            $_.Exception)
        $Logger.Error("Collection insert failed for $($CollectionItem.CollectionName): $_", "CollectionInsert")

        # Enhanced error logging
        if ($null -ne $lastMemberChangeTime) {
            $Logger.Debug("DateTime value: $($lastMemberChangeTime.ToString('yyyy-MM-dd HH:mm:ss')), Type: $($lastMemberChangeTime.GetType().FullName)", "CollectionInsert")
        }
        $Logger.Debug("Days value: $daysSinceLastChange", "CollectionInsert")

        return $false
    }
}

function Add-ApplicationToDatabase {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$CollectionItem,

        [Parameter(Mandatory = $true)]
        [System.Data.SqlClient.SqlCommand]$MonitoringCommand,

        [Parameter(Mandatory = $true)]
        [Logger]$Logger,

        [Parameter(Mandatory = $true)]
        [ErrorTracker]$ErrorTracker,

        [Parameter(Mandatory = $true)]
        [datetime]$Timestamp  # CHANGED from [string] to [datetime]
    )

    try {
        $applicationName = ConvertTo-SafeSqlString -Value $CollectionItem.ApplicationName
        $applicationPath = ConvertTo-SafeSqlString -Value $CollectionItem.FolderPath
        $packageId = ConvertTo-SafeSqlString -Value $CollectionItem."Package ID"
        $packageType = ConvertTo-SafeSqlString -Value $CollectionItem."Package Type"
        $appComments = ConvertTo-SafeSqlString -Value $CollectionItem."App_Comments"
        $appref = [System.Convert]::ToBoolean($CollectionItem."App References")

        $Logger.Debug("Inserting application: $applicationName for collection: $($CollectionItem.CollectionID)", "ApplicationInsert")

        $sqlInsert = @"
INSERT INTO [dbo].[tbl_SCCMapplications]
    ([collectionid],[collectionsiteid],[name],[path],[packageid],[packagesizeKB],[packagetype],
    [App_Comments],[appref],[eppstatus],[timestamp])
SELECT @CollectionID, @CollectionSiteID, @Name, @Path, @PackageID, @PackageSizeKB, @PackageType,
       @AppComments, @AppRef, @EPPStatus, @Timestamp
WHERE NOT EXISTS (
    SELECT [packageid] FROM [dbo].[tbl_SCCMapplications]
    WHERE [packageid] = @PackageID
        AND [collectionid] = @CollectionID
        AND [path] = @Path
)
"@

        $MonitoringCommand.Parameters.Clear()

        [void]$MonitoringCommand.Parameters.AddWithValue("@CollectionID", $CollectionItem.CollectionID)
        [void]$MonitoringCommand.Parameters.AddWithValue("@CollectionSiteID", $CollectionItem.CollectionSiteID)
        [void]$MonitoringCommand.Parameters.AddWithValue("@Name", $applicationName)
        [void]$MonitoringCommand.Parameters.AddWithValue("@Path", $applicationPath)
        [void]$MonitoringCommand.Parameters.AddWithValue("@PackageID", $packageId)
        [void]$MonitoringCommand.Parameters.AddWithValue("@PackageSizeKB", $CollectionItem."PackageSizeKB")
        [void]$MonitoringCommand.Parameters.AddWithValue("@PackageType", $packageType)
        [void]$MonitoringCommand.Parameters.AddWithValue("@AppComments", $appComments)
        [void]$MonitoringCommand.Parameters.AddWithValue("@AppRef", $appref)
        [void]$MonitoringCommand.Parameters.AddWithValue("@EPPStatus", $CollectionItem.eppstatus)

        # KEY FIX: Add timestamp as DateTime object
        $timestampParam = $MonitoringCommand.Parameters.Add("@Timestamp", [System.Data.SqlDbType]::DateTime)
        $timestampParam.Value = $Timestamp

        $MonitoringCommand.CommandText = $sqlInsert
        [void]$MonitoringCommand.ExecuteNonQuery()

        $Logger.Info("Successfully inserted application: $applicationName", "ApplicationInsert")
        return $true
    }
    catch {
        $ErrorTracker.RecordError("ApplicationInsert",
            "Failed to insert application: $($CollectionItem.ApplicationName)",
            $false,
            $_.Exception)
        $Logger.Error("Application insert failed: $_", "ApplicationInsert")
        return $false
    }
}

function Add-TaskSequenceToDatabase {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$CollectionItem,

        [Parameter(Mandatory = $true)]
        [System.Data.SqlClient.SqlCommand]$MonitoringCommand,

        [Parameter(Mandatory = $true)]
        [Logger]$Logger,

        [Parameter(Mandatory = $true)]
        [ErrorTracker]$ErrorTracker,

        [Parameter(Mandatory = $true)]
        [datetime]$Timestamp  # CHANGED from [string] to [datetime]
    )

    try {
        # Sanitize values
        $tsName = ConvertTo-SafeSqlString -Value $CollectionItem."TS Name"

        # Convert boolean
        $apptsref = [System.Convert]::ToBoolean($CollectionItem."App TS References")

        $Logger.Debug("Inserting task sequence: $tsName for collection: $($CollectionItem.CollectionID)", "TaskSequenceInsert")

        # FIXED: Use parameterized query
        $sqlInsert = @"
INSERT INTO [dbo].[tbl_SCCMtasksequences]
    ([collectionid],[collectionsiteid],[name],[apptsref],[timestamp])
SELECT @CollectionID, @CollectionSiteID, @Name, @AppTSRef, @Timestamp
WHERE NOT EXISTS (
    SELECT [collectionid],[name] FROM [dbo].[tbl_SCCMtasksequences]
    WHERE [collectionid] = @CollectionID
        AND [name] = @Name
)
"@

        # Clear existing parameters
        $MonitoringCommand.Parameters.Clear()

        # Add parameters
        [void]$MonitoringCommand.Parameters.AddWithValue("@CollectionID", $CollectionItem.CollectionID)
        [void]$MonitoringCommand.Parameters.AddWithValue("@CollectionSiteID", $CollectionItem.CollectionSiteID)
        [void]$MonitoringCommand.Parameters.AddWithValue("@Name", $tsName)
        [void]$MonitoringCommand.Parameters.AddWithValue("@AppTSRef", $apptsref)

        # KEY FIX: Add timestamp as DateTime object
        $timestampParam = $MonitoringCommand.Parameters.Add("@Timestamp", [System.Data.SqlDbType]::DateTime)
        $timestampParam.Value = $Timestamp

        $MonitoringCommand.CommandText = $sqlInsert
        [void]$MonitoringCommand.ExecuteNonQuery()

        $Logger.Info("Successfully inserted task sequence: $tsName", "TaskSequenceInsert")
        return $true
    }
    catch {
        $ErrorTracker.RecordError("TaskSequenceInsert",
            "Failed to insert task sequence: $($CollectionItem.'TS Name')",
            $false,
            $_.Exception)
        $Logger.Error("Task sequence insert failed: $_", "TaskSequenceInsert")
        return $false
    }
}

function Add-CollectionStatusToDatabase {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$CollectionId,

        [Parameter(Mandatory = $true)]
        [System.Data.SqlClient.SqlCommand]$MonitoringCommand,

        [Parameter(Mandatory = $true)]
        [Logger]$Logger,

        [Parameter(Mandatory = $true)]
        [ErrorTracker]$ErrorTracker,

        [Parameter(Mandatory = $true)]
        [datetime]$Timestamp,  # CHANGED from [string] to [datetime]

        [Parameter(Mandatory = $false)]
        [string]$Status = "Discovered",

        [Parameter(Mandatory = $false)]
        [string]$Message = "Initial discovery scan",

        [Parameter(Mandatory = $false)]
        [string]$Actions = "None"
    )

    try {
        $Logger.Debug("Inserting collection status for: $CollectionId", "StatusInsert")

        # FIXED: Use parameterized query
        $sqlInsert = @"
INSERT INTO [dbo].[tbl_SCCMCollectionStatus]
    ([collectionid],[status],[message],[actions],[timestamp])
SELECT @CollectionID, @Status, @Message, @Actions, @Timestamp
WHERE NOT EXISTS (
    SELECT [collectionid] FROM [dbo].[tbl_SCCMCollectionStatus]
    WHERE [collectionid] = @CollectionID
)
"@

        # Clean existing parameters
        $MonitoringCommand.Parameters.Clear()

        # Add parameters
        [void]$MonitoringCommand.Parameters.AddWithValue("@CollectionID", $CollectionId)
        [void]$MonitoringCommand.Parameters.AddWithValue("@Status", $Status)
        [void]$MonitoringCommand.Parameters.AddWithValue("@Message", $Message)
        [void]$MonitoringCommand.Parameters.AddWithValue("@Actions", $Actions)

        # KEY FIX: Add timestamp as DateTime object
        $timestampParam = $MonitoringCommand.Parameters.Add("@Timestamp", [System.Data.SqlDbType]::DateTime)
        $timestampParam.Value = $Timestamp

        $MonitoringCommand.CommandText = $sqlInsert
        [void]$MonitoringCommand.ExecuteNonQuery()

        $Logger.Debug("Successfully inserted collection status", "StatusInsert")
        return $true
    }
    catch {
        $ErrorTracker.RecordError("StatusInsert",
            "Failed to insert collection status for: $CollectionId",
            $false,
            $_.Exception)
        $Logger.Error("Collection status insert failed: $_", "StatusInsert")
        return $false
    }
}

function Add-TaskSequenceEPPCleanup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$TaskSequenceItem,

        [Parameter(Mandatory = $true)]
        [string]$SiteId,

        [Parameter(Mandatory = $true)]
        [System.Data.SqlClient.SqlCommand]$MonitoringCommand,

        [Parameter(Mandatory = $true)]
        [Logger]$Logger,

        [Parameter(Mandatory = $true)]
        [ErrorTracker]$ErrorTracker,

        [Parameter(Mandatory = $true)]
        [datetime]$Timestamp
    )

    try {
        # Sanitize all values
        $taskSequenceName = ConvertTo-SafeSqlString -Value $TaskSequenceItem."Task Sequence Name"
        $taskSequencePackageID = ConvertTo-SafeSqlString -Value $TaskSequenceItem."Task Sequence Package ID"
        $taskSequenceObjectPath = ConvertTo-SafeSqlString -Value $TaskSequenceItem."Task Sequence Object Path"
        $referencedPackageName = ConvertTo-SafeSqlString -Value $TaskSequenceItem."Referenced Package Name"
        $referencePackageID = ConvertTo-SafeSqlString -Value $TaskSequenceItem."ReferencePackageID"
        $collectionName = ConvertTo-SafeSqlString -Value $TaskSequenceItem."Collection Name"
        $collectionID = ConvertTo-SafeSqlString -Value $TaskSequenceItem."Collection ID"
        $collectionType = ConvertTo-SafeSqlString -Value $TaskSequenceItem."Collection Type"
        $collectionObjectPath = ConvertTo-SafeSqlString -Value $TaskSequenceItem."Collection Object Path"
        $eppStatus = ConvertTo-SafeSqlString -Value $TaskSequenceItem."EPP Status"
        $matchType = ConvertTo-SafeSqlString -Value $TaskSequenceItem."Match Type"

        # Determine collection message and status based on EPP status (NOT retired folder)
        $collectionMessage = if ($collectionName -ne "") {
            "Collection associated with task sequence referencing $eppStatus EPP package"
        } else {
            "None"
        }
        $collectionStatus = if ($collectionName -ne "") { "Discovered" } else { "None" }

        # Determine task sequence message based on EPP status
        $tsMessage = "Task sequence references $eppStatus EPP package: $referencedPackageName"
        $tsStatus = "Discovered"

        $Logger.Debug("Processing TS EPP cleanup: $taskSequenceName for SiteId: $SiteId (EPP Status: $eppStatus)", "TSEPPCleanup")

        # UPDATED: Added siteid to INSERT statement
        $sqlInsert = @"
INSERT INTO [dbo].[tbl_SCCMTaskSequencesEPPCleanup]
    ([siteid], [Task Sequence Name],[Task Sequence Package ID], [Task Sequence Path], [Referenced Package Name],
    [Reference Package ID], [Collection Name], [Collection ID], [Collection Type], [Collection Path],
    [EPP Status], [Match Type], [Timestamp], [Task Sequence Message], [Collection Message],
    [Task Sequence Status], [Collection Status], [Actions], [StatusTimestamp])
SELECT @SiteId, @TSName, @TSPackageID, @TSPath, @RefPackageName, @RefPackageID,
       @CollectionName, @CollectionID, @CollectionType, @CollectionPath,
       @EPPStatus, @MatchType, @Timestamp, @TSMessage, @CollectionMessage,
       @TSStatus, @CollectionStatus, @Actions, @StatusTimestamp
WHERE NOT EXISTS (
    SELECT 1 FROM [dbo].[tbl_SCCMTaskSequencesEPPCleanup]
    WHERE [siteid] = @SiteId
        AND [Task Sequence Name] = @TSName
        AND [Referenced Package Name] = @RefPackageName
        AND [Collection ID] = @CollectionID
        AND [EPP Status] = @EPPStatus
        AND [Match Type] = @MatchType
)
"@

        # Clear existing parameters
        $MonitoringCommand.Parameters.Clear()

        # ADDED: SiteId parameter
        [void]$MonitoringCommand.Parameters.AddWithValue("@SiteId", $SiteId)

        # Add other parameters
        [void]$MonitoringCommand.Parameters.AddWithValue("@TSName", $taskSequenceName)
        [void]$MonitoringCommand.Parameters.AddWithValue("@TSPackageID", $taskSequencePackageID)
        [void]$MonitoringCommand.Parameters.AddWithValue("@TSPath", $taskSequenceObjectPath)
        [void]$MonitoringCommand.Parameters.AddWithValue("@RefPackageName", $referencedPackageName)
        [void]$MonitoringCommand.Parameters.AddWithValue("@RefPackageID", $referencePackageID)
        [void]$MonitoringCommand.Parameters.AddWithValue("@CollectionName", $collectionName)
        [void]$MonitoringCommand.Parameters.AddWithValue("@CollectionID", $collectionID)
        [void]$MonitoringCommand.Parameters.AddWithValue("@CollectionType", $collectionType)
        [void]$MonitoringCommand.Parameters.AddWithValue("@CollectionPath", $collectionObjectPath)
        [void]$MonitoringCommand.Parameters.AddWithValue("@EPPStatus", $eppStatus)
        [void]$MonitoringCommand.Parameters.AddWithValue("@MatchType", $matchType)

        # Add timestamp as DateTime object
        $timestampParam = $MonitoringCommand.Parameters.Add("@Timestamp", [System.Data.SqlDbType]::DateTime)
        $timestampParam.Value = $Timestamp

        # Use dynamic messages based on EPP status
        [void]$MonitoringCommand.Parameters.AddWithValue("@TSMessage", $tsMessage)
        [void]$MonitoringCommand.Parameters.AddWithValue("@CollectionMessage", $collectionMessage)
        [void]$MonitoringCommand.Parameters.AddWithValue("@TSStatus", $tsStatus)
        [void]$MonitoringCommand.Parameters.AddWithValue("@CollectionStatus", $collectionStatus)
        [void]$MonitoringCommand.Parameters.AddWithValue("@Actions", "Review for cleanup")

        # Add status timestamp as DateTime object
        $statusTimestampParam = $MonitoringCommand.Parameters.Add("@StatusTimestamp", [System.Data.SqlDbType]::DateTime)
        $statusTimestampParam.Value = $Timestamp

        $MonitoringCommand.CommandText = $sqlInsert
        [void]$MonitoringCommand.ExecuteNonQuery()

        $Logger.Info("Successfully processed TS EPP cleanup: $taskSequenceName (SiteId: $SiteId, EPP: $eppStatus)", "TSEPPCleanup")
        return $true
    }
    catch {
        $ErrorTracker.RecordError("TSEPPCleanup",
            "Failed to process TS EPP cleanup: $($TaskSequenceItem.'Task Sequence Name')",
            $false,
            $_.Exception)
        $Logger.Error("TS EPP cleanup failed: $_", "TSEPPCleanup")
        return $false
    }
}

function Get-ApplicationReferences {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [DatabaseConnection]$SCCMConnection,

        [Parameter(Mandatory = $true)]
        [string]$ApplicationName,

        [Parameter(Mandatory = $true)]
        [Logger]$Logger
    )

    try {
        $Logger.Debug("Checking application references for: $ApplicationName", "AppRefChecker")

        $refQuery = @"
SET NOCOUNT ON

-- Check if application is referenced by other applications
;WITH AppReferences AS (
    SELECT DISTINCT
        ParentApp.DisplayName AS ParentAppName,
        ChildApp.DisplayName AS ChildAppName,
        'Application Dependency' AS ReferenceType
    FROM fn_ListApplicationCIs(1033) ParentApp
    LEFT JOIN vSMS_AppRelation_Flat R
        ON R.FromApplicationCIID = ParentApp.CI_ID
    LEFT JOIN fn_ListApplicationCIs_List(1033) ChildApp
        ON ChildApp.CI_ID = R.ToApplicationCIID
        AND ChildApp.IsLatest = 1
    WHERE ParentApp.IsLatest = 1
        AND ChildApp.DisplayName = @AppName
),
-- Check if application is referenced by task sequences
TSReferences AS (
    SELECT DISTINCT
        TS.Name AS TaskSequenceName,
        app.DisplayName AS ApplicationName,
        'Task Sequence Reference' AS ReferenceType
    FROM fn_ListLatestApplicationCIs(1033) app
    INNER JOIN v_TaskSequenceAppReferencesInfo TSApp
        ON app.ModelName = TSApp.RefAppModelName
    INNER JOIN v_TaskSequencePackage TS
        ON TS.PackageID = TSApp.PackageID
    WHERE app.DisplayName = @AppName
),
-- Check if application has active deployments
DeploymentReferences AS (
    SELECT DISTINCT
        APP.DisplayName AS ApplicationName,
        C.Name AS CollectionName,
        'Active Deployment' AS ReferenceType
    FROM fn_ListApplicationCIs(1033) APP
    JOIN fn_ListDeploymentTypeCIs(1033) DT
        ON DT.AppModelName = APP.ModelName
        AND DT.IsLatest = 1
    LEFT JOIN v_CIAssignmentToCI CIACI
        ON CIACI.CI_ID = APP.CI_ID
    LEFT JOIN v_CIAssignment CIA
        ON CIACI.AssignmentID = CIA.AssignmentID
    LEFT JOIN v_Collection C
        ON C.CollectionID = CIA.CollectionID
    WHERE APP.IsLatest = 1
        AND APP.DisplayName = @AppName
        AND C.Name IS NOT NULL
)
-- Combine all references
SELECT * FROM AppReferences
UNION ALL
SELECT TaskSequenceName AS ParentAppName, ApplicationName AS ChildAppName, ReferenceType FROM TSReferences
UNION ALL
SELECT CollectionName AS ParentAppName, ApplicationName AS ChildAppName, ReferenceType FROM DeploymentReferences
"@

        $dataset = New-Object System.Data.DataSet
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($refQuery, $SCCMConnection.Connection)
        $adapter.SelectCommand.CommandTimeout = 300

        # Add parameter
        [void]$adapter.SelectCommand.Parameters.AddWithValue("@AppName", $ApplicationName)

        [void]$adapter.Fill($dataset)

        $references = $dataset.Tables[0]

        $isOrphaned = ($references.Rows.Count -eq 0)

        $Logger.Debug("Application '$ApplicationName' - Orphaned: $isOrphaned, References: $($references.Rows.Count)", "AppRefChecker")

        return @{
            IsOrphaned = $isOrphaned
            ReferenceCount = $references.Rows.Count
            References = $references
        }
    }
    catch {
        $Logger.Warning("Failed to check application references for '$ApplicationName': $_", "AppRefChecker")
        return @{
            IsOrphaned = $false
            ReferenceCount = -1
            References = $null
        }
    }
}

function New-OrphanedObjectDetectionQuery {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ObjectType,  # 'DeviceCollection', 'UserCollection', 'TaskSequence'

        [Parameter(Mandatory = $true)]
        [Logger]$Logger
    )

    try {
        $Logger.Info("Building orphaned object detection query for: $ObjectType", "OrphanQueryBuilder")

        $baseQuery = @"
SET NOCOUNT ON

-- CTE for folder hierarchy
;WITH FolderHierarchy AS (
    SELECT ContainerNodeID,
           CONVERT(nvarchar(max), Name + N'\') AS [Path],
           Name
    FROM dbo.Folders
    WHERE ParentContainerNodeID = 0

    UNION ALL

    SELECT fld.ContainerNodeID,
           cte.Path + fld.Name + N'\',
           fld.Name
    FROM dbo.Folders fld
    INNER JOIN FolderHierarchy cte
        ON fld.ParentContainerNodeID = cte.ContainerNodeID
)
"@

        switch ($ObjectType) {
            'DeviceCollection' {
                $query = $baseQuery + @"
,
-- Get all device collections in staging folders
StagingCollections AS (
    SELECT DISTINCT
        col.CollectionID,
        col.SiteID AS CollectionSiteID,
        col.CollectionName,
        ISNULL(hry.[Path], '/') AS FolderPath,
        col.MemberCount,
        col.LastMemberChangeTime,
        DATEDIFF(day, col.LastMemberChangeTime, GETDATE()) AS DaysSinceLastChange,
        col.LimitToCollectionID,
        col.LimitToCollectionName
    FROM dbo.v_Collections col WITH (NOLOCK)
    LEFT JOIN dbo.vFolderMembers mem WITH (NOLOCK)
        ON col.SiteID = mem.InstanceKey
    LEFT JOIN FolderHierarchy hry
        ON mem.ContainerNodeID = hry.ContainerNodeID
    WHERE col.CollectionType = 2  -- Device Collections
        AND mem.ObjectType = '5000'  -- Device Collection object type
        AND hry.[Path] LIKE '%Staging%'
        AND col.PowerConfigsCount = '0'
        AND col.SiteID NOT IN (
            SELECT CollectionID
            FROM vSMS_ClientSettingsAssignments WITH (NOLOCK)
        )
)
SELECT * FROM StagingCollections
ORDER BY CollectionName
"@
            }
            'UserCollection' {
                $query = $baseQuery + @"
,
-- Get all user collections in staging folders
StagingCollections AS (
    SELECT DISTINCT
        col.CollectionID,
        col.SiteID AS CollectionSiteID,
        col.CollectionName,
        ISNULL(hry.[Path], '/') AS FolderPath,
        col.MemberCount,
        col.LastMemberChangeTime,
        DATEDIFF(day, col.LastMemberChangeTime, GETDATE()) AS DaysSinceLastChange,
        col.LimitToCollectionID,
        col.LimitToCollectionName
    FROM dbo.v_Collections col WITH (NOLOCK)
    LEFT JOIN dbo.vFolderMembers mem WITH (NOLOCK)
        ON col.SiteID = mem.InstanceKey
    LEFT JOIN FolderHierarchy hry
        ON mem.ContainerNodeID = hry.ContainerNodeID
    WHERE col.CollectionType = 1  -- User Collections
        AND mem.ObjectType = '5001'  -- User Collection object type
        AND hry.[Path] LIKE '%Staging%'
        AND col.PowerConfigsCount = '0'
        AND col.SiteID NOT IN (
            SELECT CollectionID
            FROM vSMS_ClientSettingsAssignments WITH (NOLOCK)
        )
)
SELECT * FROM StagingCollections
ORDER BY CollectionName
"@
            }
            'TaskSequence' {
                $query = $baseQuery + @"
,
-- Get all task sequences in staging folders
StagingTaskSequences AS (
    SELECT DISTINCT
        tsp.PackageID,
        tsp.Name AS TaskSequenceName,
        ISNULL(hry.[Path], '/') AS FolderPath,
        tsp.SourceDate,
        DATEDIFF(day, tsp.SourceDate, GETDATE()) AS DaysSinceLastChange,
        tsp.Description
    FROM v_TaskSequencePackage tsp WITH (NOLOCK)
    LEFT JOIN vFolderMembers mem WITH (NOLOCK)
        ON tsp.PackageID = mem.InstanceKey
    LEFT JOIN FolderHierarchy hry
        ON mem.ContainerNodeID = hry.ContainerNodeID
    WHERE hry.[Path] LIKE '%Staging%'
)
SELECT * FROM StagingTaskSequences
ORDER BY TaskSequenceName
"@
            }
            default {
                throw "Invalid ObjectType: $ObjectType"
            }
        }

        $Logger.Info("Query built successfully for $ObjectType", "OrphanQueryBuilder")
        return $query
    }
    catch {
        $Logger.Error("Failed to build orphaned object detection query: $_", "OrphanQueryBuilder")
        throw
    }
}

#endregion

#region Staging Orphaned Objects Processing Orchestrator with Retired Object Exclusion
function Invoke-StagingOrphanedObjectsProcessing {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [DatabaseConnection]$SCCMConnection,

        [Parameter(Mandatory = $true)]
        [string]$SiteId,

        [Parameter(Mandatory = $true)]
        [System.Data.SqlClient.SqlCommand]$MonitoringCommand,

        [Parameter(Mandatory = $true)]
        [Logger]$Logger,

        [Parameter(Mandatory = $true)]
        [ErrorTracker]$ErrorTracker,

        [Parameter(Mandatory = $true)]
        [ResourceManager]$ResourceManager,

        [Parameter(Mandatory = $true)]
        [datetime]$Timestamp,

        # NEW PARAMETER: Retired folder data to exclude from orphaned scan
        [Parameter(Mandatory = $false)]
        [System.Data.DataTable]$RetiredFolderData = $null,

        [Parameter(Mandatory = $false)]
        [ProgressTracker]$Progress,

        [Parameter(Mandatory = $false)]
        [int]$SiteIndex = 1,

        [Parameter(Mandatory = $false)]
        [int]$TotalSites = 1
    )

    try {
        $Logger.Info("Starting staging orphaned objects processing for site", "StagingOrphanProcessor")

        # NEW: Build exclusion lists from retired folder data
        $retiredCollectionIds = @()
        $retiredTaskSequenceIds = @()

        if ($null -ne $RetiredFolderData -and $RetiredFolderData.Rows.Count -gt 0) {
            # Extract retired collection IDs
            $retiredCollectionIds = $RetiredFolderData.Rows |
                Where-Object {
                    ($_.ObjectType -eq 'UserCollection' -or $_.ObjectType -eq 'DeviceCollection') -and
                    -not [string]::IsNullOrWhiteSpace($_.CollectionSiteID)
                } |
                ForEach-Object { $_.CollectionSiteID.ToString().Trim() }

            # Extract retired task sequence package IDs
            $retiredTaskSequenceIds = $RetiredFolderData.Rows |
                Where-Object {
                    $_.ObjectType -eq 'TaskSequence' -and
                    -not [string]::IsNullOrWhiteSpace($_.PackageID)
                } |
                ForEach-Object { $_.PackageID.ToString().Trim() }

            $Logger.Info("Exclusion lists built - Collections: $($retiredCollectionIds.Count), Task Sequences: $($retiredTaskSequenceIds.Count)", "StagingOrphanProcessor")
        }

        $totalProcessed = 0
        $totalErrors = 0
        $totalExcluded = 0

        # Scan Device Collections
        if ($Progress) {
            $Progress.UpdateProgress("ProcessSites",
                "Site $SiteIndex/$TotalSites - Scanning staging device collections for orphans",
                (($SiteIndex - 1) / $TotalSites * 100) + 5)
        }

        $deviceCollectionResult = Invoke-StagingOrphanedCollectionScan `
            -SCCMConnection $SCCMConnection `
            -CollectionType 'Device' `
            -Logger $Logger `
            -ErrorTracker $ErrorTracker `
            -ResourceManager $ResourceManager `
            -Progress $Progress `
            -SiteIndex $SiteIndex `
            -TotalSites $TotalSites

        # Process orphaned device collections (with exclusion)
        if ($deviceCollectionResult.Success -and $deviceCollectionResult.OrphanedCollections.Count -gt 0) {
            $Logger.Info("Processing $($deviceCollectionResult.OrphanedCollections.Count) orphaned device collections", "StagingOrphanProcessor")

            foreach ($orphanedCollection in $deviceCollectionResult.OrphanedCollections) {
                $collectionSiteId = $orphanedCollection.Collection.CollectionSiteID.ToString().Trim()

                # NEW: Check if this collection is already in retired folders
                if ($collectionSiteId -in $retiredCollectionIds) {
                    $totalExcluded++
                    $Logger.Debug("Excluded device collection (already in retired folder): $($orphanedCollection.Collection.CollectionName)", "StagingOrphanProcessor")
                    continue
                }

                $success = Add-OrphanedObjectToDatabase `
                    -OrphanedObject $orphanedCollection `
                    -SiteId $SiteId `
                    -ObjectType 'Collection' `
                    -MonitoringCommand $MonitoringCommand `
                    -Logger $Logger `
                    -ErrorTracker $ErrorTracker `
                    -Timestamp $Timestamp

                if ($success) {
                    $totalProcessed++
                } else {
                    $totalErrors++
                }
            }
        }

        # Scan User Collections
        if ($Progress) {
            $Progress.UpdateProgress("ProcessSites",
                "Site $SiteIndex/$TotalSites - Scanning staging user collections for orphans",
                (($SiteIndex - 1) / $TotalSites * 100) + 6)
        }

        $userCollectionResult = Invoke-StagingOrphanedCollectionScan `
            -SCCMConnection $SCCMConnection `
            -CollectionType 'User' `
            -Logger $Logger `
            -ErrorTracker $ErrorTracker `
            -ResourceManager $ResourceManager `
            -Progress $Progress `
            -SiteIndex $SiteIndex `
            -TotalSites $TotalSites

        # Process orphaned user collections (with exclusion)
        if ($userCollectionResult.Success -and $userCollectionResult.OrphanedCollections.Count -gt 0) {
            $Logger.Info("Processing $($userCollectionResult.OrphanedCollections.Count) orphaned user collections", "StagingOrphanProcessor")

            foreach ($orphanedCollection in $userCollectionResult.OrphanedCollections) {
                $collectionSiteId = $orphanedCollection.Collection.CollectionSiteID.ToString().Trim()

                # NEW: Check if this collection is already in retired folders
                if ($collectionSiteId -in $retiredCollectionIds) {
                    $totalExcluded++
                    $Logger.Debug("Excluded user collection (already in retired folder): $($orphanedCollection.Collection.CollectionName)", "StagingOrphanProcessor")
                    continue
                }

                $success = Add-OrphanedObjectToDatabase `
                    -OrphanedObject $orphanedCollection `
                    -SiteId $SiteId `
                    -ObjectType 'Collection' `
                    -MonitoringCommand $MonitoringCommand `
                    -Logger $Logger `
                    -ErrorTracker $ErrorTracker `
                    -Timestamp $Timestamp

                if ($success) {
                    $totalProcessed++
                } else {
                    $totalErrors++
                }
            }
        }

        # Scan Task Sequences
        if ($Progress) {
            $Progress.UpdateProgress("ProcessSites",
                "Site $SiteIndex/$TotalSites - Scanning staging task sequences for orphans",
                (($SiteIndex - 1) / $TotalSites * 100) + 7)
        }

        $taskSequenceResult = Invoke-StagingOrphanedTaskSequenceScan `
            -SCCMConnection $SCCMConnection `
            -Logger $Logger `
            -ErrorTracker $ErrorTracker `
            -ResourceManager $ResourceManager `
            -Progress $Progress `
            -SiteIndex $SiteIndex `
            -TotalSites $TotalSites

        # Process orphaned task sequences (with exclusion)
        if ($taskSequenceResult.Success -and $taskSequenceResult.OrphanedTaskSequences.Count -gt 0) {
            $Logger.Info("Processing $($taskSequenceResult.OrphanedTaskSequences.Count) orphaned task sequences", "StagingOrphanProcessor")

            foreach ($orphanedTS in $taskSequenceResult.OrphanedTaskSequences) {
                $packageId = $orphanedTS.TaskSequence.PackageID.ToString().Trim()

                # NEW: Check if this task sequence is already in retired folders
                if ($packageId -in $retiredTaskSequenceIds) {
                    $totalExcluded++
                    $Logger.Debug("Excluded task sequence (already in retired folder): $($orphanedTS.TaskSequence.TaskSequenceName)", "StagingOrphanProcessor")
                    continue
                }

                $success = Add-OrphanedObjectToDatabase `
                    -OrphanedObject $orphanedTS `
                    -SiteId $SiteId `
                    -ObjectType 'TaskSequence' `
                    -MonitoringCommand $MonitoringCommand `
                    -Logger $Logger `
                    -ErrorTracker $ErrorTracker `
                    -Timestamp $Timestamp

                if ($success) {
                    $totalProcessed++
                } else {
                    $totalErrors++
                }
            }
        }

        $Logger.Info("Staging orphaned objects processing completed - Processed: $totalProcessed, Excluded: $totalExcluded, Errors: $totalErrors", "StagingOrphanProcessor")

        return @{
            Success = $true
            ProcessedCount = $totalProcessed
            ExcludedCount = $totalExcluded
            ErrorCount = $totalErrors
            DeviceCollections = $deviceCollectionResult.OrphanedCollections.Count
            UserCollections = $userCollectionResult.OrphanedCollections.Count
            TaskSequences = $taskSequenceResult.OrphanedTaskSequences.Count
        }
    } catch {
        $ErrorTracker.RecordError("StagingOrphanProcessor",
            "Fatal error during staging orphaned objects processing",
            $true,
            $_.Exception)
        $Logger.Fatal("Staging orphaned objects processing failed: $_", "StagingOrphanProcessor")

        return @{
            Success = $false
            ProcessedCount = 0
            ExcludedCount = 0
            ErrorCount = 0
            DeviceCollections = 0
            UserCollections = 0
            TaskSequences = 0
        }
    }
    }
#endregion

#region Task Sequence Query Execution
function Get-TaskSequenceData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [DatabaseConnection]$SCCMConnection,

        [Parameter(Mandatory = $true)]
        [Logger]$Logger,

        [Parameter(Mandatory = $true)]
        [ErrorTracker]$ErrorTracker,

        [Parameter(Mandatory = $true)]
        [ResourceManager]$ResourceManager
    )

    try {
        $Logger.Info("Executing Task Sequence query", "TaskSequenceQuery")

        # SCCM query to get Task Sequences, referenced packages, and collection names
        $sccmQuery = @"
            -- CTE for folder hierarchy to build object paths
            WITH FolderHierarchy AS (
                SELECT ContainerNodeID,
                    CONVERT(nvarchar(max), Name + N'\') AS [Path]
                FROM dbo.Folders
                WHERE ParentContainerNodeID = 0

                UNION ALL

                SELECT fld.ContainerNodeID,
                    cte.Path + fld.Name + N'\'
                FROM dbo.Folders fld
                INNER JOIN FolderHierarchy cte
                    ON fld.ParentContainerNodeID = cte.ContainerNodeID
            )
            SELECT
                tsp.PackageID AS [Task Sequence Package ID],
                tsp.Name AS [Task Sequence Name],
                CASE WHEN ts_hry.[Path] IS NULL
                    THEN '\'
                    ELSE ts_hry.[Path]
                END AS [Task Sequence Object Path],
                tsri.ReferencePackageID,
                pkg.Name AS [Referenced Package Name],
                coll.Name AS [Collection Name],
                coll.CollectionID,
                CASE WHEN coll.Name IS NULL THEN NULL
                    WHEN coll.CollectionType = 1 THEN 'UserCollection'
                    ELSE 'DeviceCollection'
                END AS [Collection Type],
                CASE WHEN coll_hry.[Path] IS NULL
                    THEN '\'
                    ELSE coll_hry.[Path]
                END AS [Collection Object Path]
            FROM v_TaskSequencePackage tsp
            JOIN v_TaskSequenceReferencesInfo tsri ON tsp.PackageID = tsri.PackageID
            JOIN v_Package pkg ON tsri.ReferencePackageID = pkg.PackageID
            LEFT JOIN v_Advertisement adv ON tsp.PackageID = adv.PackageID
            LEFT JOIN v_Collection coll ON adv.CollectionID = coll.CollectionID
            LEFT JOIN vFolderMembers ts_mem ON tsp.PackageID = ts_mem.InstanceKey
            LEFT JOIN FolderHierarchy ts_hry ON ts_mem.ContainerNodeID = ts_hry.ContainerNodeID
            LEFT JOIN vFolderMembers coll_mem ON coll.CollectionID = coll_mem.InstanceKey
            LEFT JOIN FolderHierarchy coll_hry ON coll_mem.ContainerNodeID = coll_hry.ContainerNodeID
"@

        $dataset = New-Object System.Data.DataSet
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($sccmQuery, $SCCMConnection.Connection)
        $adapter.SelectCommand.CommandTimeout = 40000

        $ResourceManager.Register($adapter, "TaskSequenceAdapter")

        [void]$adapter.Fill($dataset)

        $dataTable = New-Object System.Data.DataTable
        $stagingTable = New-Object System.Data.DataTable
        $dataTable = $dataset.Tables[0]
        $dataTable.Merge($stagingTable)

        $Logger.Info("Task Sequence query completed. Rows returned: $($dataTable.Rows.Count)", "TaskSequenceQuery")

        if ($null -eq $dataTable -or $dataTable.Rows.Count -eq 0) {
            $Logger.Warning("Task Sequence query returned no data", "TaskSequenceQuery")
        }

        return @{
            Success = $true
            Data = $dataTable
        }
    } catch {
        $ErrorTracker.RecordError("TaskSequenceQuery",
            "Failed to execute Task Sequence query",
            $false,
            $_.Exception)
        $Logger.Error("Task Sequence query failed: $_", "TaskSequenceQuery")

        return @{
            Success = $false
            Data = $null
        }
    }
}
#endregion

#region EPP Data Retrieval
function Get-EPPPackageData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [DatabaseConnection]$SPDWConnection,

        [Parameter(Mandatory = $true)]
        [Logger]$Logger,

        [Parameter(Mandatory = $true)]
        [ErrorTracker]$ErrorTracker,

        [Parameter(Mandatory = $true)]
        [ResourceManager]$ResourceManager,

        [Parameter(Mandatory = $false)]
        [int]$MaxRetries = 3,

        [Parameter(Mandatory = $false)]
        [int]$RetryDelaySeconds = 10
    )

    $attempt = 0

    while ($attempt -lt $MaxRetries) {
        $attempt++

        try {
            $Logger.Info("Retrieving EPP package data (Attempt $attempt of $MaxRetries)", "EPPQuery")

            # Reconnect if session is no longer valid
            if (-not $SPDWConnection.IsConnected -or
                $SPDWConnection.Connection.State -ne [System.Data.ConnectionState]::Open) {

                $Logger.Warning("SPDW connection is not open - attempting reconnect (Attempt $attempt)", "EPPQuery")
                $SPDWConnection.Disconnect($Logger)

                if (-not $SPDWConnection.Connect($Logger, $ErrorTracker)) {
                    throw "Failed to reconnect to SPDW database on attempt $attempt"
                }

                $Logger.Info("SPDW reconnection successful", "EPPQuery")
            }

            # Validate connection with lightweight test before executing main query
            try {
                $testCmd = $SPDWConnection.Connection.CreateCommand()
                $testCmd.CommandText = "SELECT 1"
                $testCmd.CommandTimeout = 10
                [void]$testCmd.ExecuteScalar()
                $testCmd.Dispose()
                $Logger.Debug("SPDW connection health check passed", "EPPQuery")
            }
            catch {
                $Logger.Warning("SPDW connection health check failed - forcing reconnect: $_", "EPPQuery")
                $SPDWConnection.Disconnect($Logger)

                if (-not $SPDWConnection.Connect($Logger, $ErrorTracker)) {
                    throw "Failed to reconnect to SPDW database after health check failure on attempt $attempt"
                }
            }

            $eepsqlquery = "SELECT [EPP Status],[Package Name],[Package Short Name] FROM [$($SPDWConnection.DatabaseName)].[dbo].[spdw_vwSFSoftwarePackages]"

            $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($eepsqlquery, $SPDWConnection.Connection)
            $dataTable = New-Object System.Data.DataTable
            $adapter.SelectCommand.CommandTimeout = 40000

            $ResourceManager.Register($adapter, "EPPAdapter-Attempt$attempt")

            [void]$adapter.Fill($dataTable)

            $Logger.Info("EPP query completed successfully on attempt $attempt. Rows returned: $($dataTable.Rows.Count)", "EPPQuery")

            return @{
                Success = $true
                Data = $dataTable
            }
        }
        catch {
            $isKillState = $_.Exception.Message -match "kill state" -or
                $_.Exception.Message -match "severe error" -or
                $_.Exception.Message -match "connection.*broken" -or
                $_.Exception.Message -match "transport-level error"

            if ($attempt -lt $MaxRetries) {
                if ($isKillState) {
                    $Logger.Warning("SPDW session killed - will retry after ${RetryDelaySeconds}s (Attempt $attempt/$MaxRetries): $($_.Exception.Message)", "EPPQuery")
                } else {
                    $Logger.Warning("EPP query failed - will retry after ${RetryDelaySeconds}s (Attempt $attempt/$MaxRetries): $($_.Exception.Message)", "EPPQuery")
                }

                # Force disconnect before retry to ensure clean state
                try {
                    $SPDWConnection.Disconnect($Logger)
                } catch {
                    $Logger.Debug("Disconnect during retry cleanup failed (non-critical): $_", "EPPQuery")
                }

                Start-Sleep -Seconds $RetryDelaySeconds
            }
            else {
                $Logger.Error("EPP query failed after $MaxRetries attempts. Last error: $($_.Exception.Message)", "EPPQuery")
                $ErrorTracker.RecordError("EPPQuery",
                    "Failed to retrieve EPP package data after $MaxRetries attempts",
                    $false,
                    $_.Exception)

                return @{
                    Success = $false
                    Data = $null
                }
            }
        }
    }

    # Should never reach here but safety net
    return @{
        Success = $false
        Data = $null
    }
}
#endregion

#region Collection Detection Query Execution with Progress
function Invoke-CollectionDetectionQuery {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [DatabaseConnection]$SCCMConnection,

        [Parameter(Mandatory = $true)]
        [string]$AdaptivaSqlDbName,

        [Parameter(Mandatory = $true)]
        [Logger]$Logger,

        [Parameter(Mandatory = $true)]
        [ErrorTracker]$ErrorTracker,

        [Parameter(Mandatory = $true)]
        [ResourceManager]$ResourceManager,

        [Parameter(Mandatory = $false)]
        [ProgressTracker]$Progress,

        [Parameter(Mandatory = $false)]
        [int]$SiteIndex = 1,

        [Parameter(Mandatory = $false)]
        [int]$TotalSites = 1
    )

    try
    {
        $Logger.Info("Executing Collection Detection query", "CollectionDetectionQuery")

        # NEW: Setup InfoMessage event handler for SQL progress messages
        $progressMessages = [System.Collections.ArrayList]::new()

        $infoMessageHandler = {
            param($connection, $infoMessage)
            $message = $infoMessage.Message

            # Parse progress messages in format: "PROGRESS:percent:description"
            if ($message -match "PROGRESS:(\d+):(.+)") {
                $percent = [int]$Matches[1]
                $description = $Matches[2]

                # Calculate overall progress including site index
                $siteBaseProgress = (($SiteIndex - 1) / $TotalSites * 100)
                $queryProgress = ($percent / 100) * 60  # Query is 60% of site processing
                $totalProgress = $siteBaseProgress + $queryProgress

                if ($Progress) {
                    $Progress.UpdateProgress("ProcessSites",
                        "Site $SiteIndex/$TotalSites - $description",
                        $totalProgress)
                }

                $Logger.Debug("SQL Progress: $percent% - $description", "CollectionDetectionQuery")
                [void]$progressMessages.Add("[$percent%] $description")
            }
        }

        # NEW: Attach event handler to connection
        $SCCMConnection.Connection.add_InfoMessage($infoMessageHandler)

        try
        {
            # Initialize query parts array
            $queryParts = @()

            # Part 1: Initial setup and compliance table
            if ($Progress) {
                $Progress.UpdateProgress("ProcessSites", "Site $SiteIndex/$TotalSites - Part 1/20: Creating temporary tables",
                    (($SiteIndex - 1) / $TotalSites * 100) + 2)
            }
            $queryParts += @"
                SET ANSI_WARNINGS OFF
                SET NOCOUNT ON
                RAISERROR('PROGRESS:2:Creating temporary tables', 0, 1) WITH NOWAIT

                CREATE TABLE #ComplianceTable (CI_ID INT, totaldevicecnt int, compliancecnt int, percompliance varchar(50))

                CREATE TABLE #CollectionHierarchy (ModelID nvarchar(50), CollectionID int, CollectionSiteID nvarchar(50), [Adaptiva Client Settings Collection Reference] bit, [Adaptiva Policy Collection Reference] bit, [Adaptiva Health Check Collection Reference] bit,
                    LastMemberChangeTime datetime, [Days Since Last Member Change] int,
                    CollectionName nvarchar (500), CollectionPath nvarchar(500), MemberCount int, CollectionType nvarchar(100), DeploymentType nvarchar(100),
                    baselinetotaldevicecnt nvarchar(50), compliancecnt nvarchar(50), baselinecompliance nvarchar(50), baselinecollectionrefcnt nvarchar(50), baselineconfigitemscount nvarchar(50),
                    baselineconfigitemreference nvarchar(50),baselinecollectionreference nvarchar(50), ApplicationName nvarchar(100), LimitingCollection nvarchar(500), ReferencedCollection nvarchar(max))
"@

            # Part 2: Staging temp table and appObjects table creation
            if ($Progress) {
                $Progress.UpdateProgress("ProcessSites", "Site $SiteIndex/$TotalSites - Part 2/20: Creating staging tables",
                    (($SiteIndex - 1) / $TotalSites * 100) + 4)
            }
            $queryParts += @"
                RAISERROR('PROGRESS:6:Creating staging and application object tables', 0, 1) WITH NOWAIT
                CREATE TABLE #CollectionHierarchyStaging (ModelID nvarchar(50), CollectionID int, CollectionSiteID nvarchar(50), [Adaptiva Client Settings Collection Reference] bit, [Adaptiva Policy Collection Reference] bit, [Adaptiva Health Check Collection Reference] bit,
                    LastMemberChangeTime datetime, [Days Since Last Member Change] int,
                    CollectionName nvarchar (500), CollectionPath nvarchar(500), MemberCount int, CollectionType nvarchar(100), DeploymentType nvarchar(100),
                    baselinetotaldevicecnt nvarchar(50), compliancecnt nvarchar(50), baselinecompliance nvarchar(50), baselinecollectionrefcnt nvarchar(50), baselineconfigitemscount nvarchar(50),
                    baselineconfigitemreference nvarchar(50),baselinecollectionreference nvarchar(50), ApplicationName nvarchar(100), LimitingCollection nvarchar(500), ReferencedCollection nvarchar(max))

                CREATE TABLE #appObjects ([DP Group] nvarchar(20), [Package Name] nvarchar(50), [Package ID] nvarchar(50), [ObjectPath] nvarchar(250), targeteddpcount int, numberinstalled int, numberinprogress int, numbererrors int, [Package Type] nvarchar(100))
"@

            # Part 3: Populate appObjects table
            if ($Progress) {
                $Progress.UpdateProgress("ProcessSites", "Site $SiteIndex/$TotalSites - Part 3/20: Loading application objects",
                    (($SiteIndex - 1) / $TotalSites * 100) + 6)
            }
            $queryParts += @"
                RAISERROR('PROGRESS:10:Loading application objects and distribution points', 0, 1) WITH NOWAIT
                INSERT INTO #appObjects
                SELECT * FROM (
                    SELECT DISTINCT dpgr.[NAME] as [DP Group],
                        pk.[NAME] [Package Name],
                        dgp.pkgid [Package ID],
                        app.[ObjectPath],
                        dpcn.targeteddpcount,
                        dpcn.numberinstalled,
                        dpcn.numberinprogress,
                        dpcn.numbererrors,
                        CASE
                            WHEN pk.packagetype = 0 THEN 'Software Distribution Package'
                            WHEN pk.packagetype = 3 THEN 'Driver Package'
                            WHEN pk.packagetype = 4 THEN 'Task Sequence Package'
                            WHEN pk.packagetype = 5 THEN 'Software Update Package'
                            WHEN pk.packagetype = 6 THEN 'Device Setting Package'
                            WHEN pk.packagetype = 7 THEN 'Virtual Package'
                            WHEN pk.packagetype = 8 THEN 'Application'
                            WHEN pk.packagetype = 257 THEN 'Image Package'
                            WHEN pk.packagetype = 258 THEN 'Boot Image Package'
                            WHEN pk.packagetype = 259 THEN 'Operating System Install Package'
                            ELSE 'Unknown'
                        END AS 'Package Type'
                    FROM vSMS_DpGroupInfo dpgr
                    INNER JOIN v_dpgroupPackages dgp ON dgp.groupid = dpgr.groupid
                    LEFT JOIN v_package pk ON pk.packageid = dgp.pkgid
                    LEFT JOIN fn_ListApplicationCIs_List(1031) app ON pk.[NAME] = app.[DisplayName]
                    LEFT JOIN v_dpgroupcontentdetails dpcn ON dpcn.groupid = dpgr.groupid AND dpcn.pkgid = pk.packageid
                ) a
                WHERE a.[ObjectPath] IS NOT NULL
                DECLARE @AppObjectCount INT = (SELECT COUNT(*) FROM #appObjects)
                RAISERROR('PROGRESS:12:Loaded %d application objects', 0, 1, @AppObjectCount) WITH NOWAIT
"@

            # Part 4: Populate ComplianceTable
            if ($Progress) {
                $Progress.UpdateProgress("ProcessSites", "Site $SiteIndex/$TotalSites - Part 4/20: Calculating compliance data",
                    (($SiteIndex - 1) / $TotalSites * 100) + 8)
            }
            $queryParts += @"
                RAISERROR('PROGRESS:14:Calculating compliance metrics for configuration items', 0, 1) WITH NOWAIT
                INSERT INTO #ComplianceTable
                SELECT CI.CI_ID, COUNT(distinct VRS.Netbios_Name0) as [totaldevicecnt], clt.[compliancecnt] as [compliancecnt],
                    FLOOR((CAST(clt.[compliancecnt] as numeric) / CAST(COUNT(distinct VRS.Netbios_Name0) as numeric)) * 100) as [perccompliance]
                FROM v_R_System VRS
                RIGHT JOIN v_FullCollectionMembership_Valid FM on VRS.ResourceID= FM.ResourceID
                RIGHT JOIN fn_ListCI_ComplianceState(1033) CI on VRS.ResourceID= CI.ResourceID
                LEFT JOIN (select CI.CI_ID, COUNT(distinct VRS.Netbios_Name0) as [compliancecnt]
                    from v_R_System VRS
                    RIGHT JOIN v_FullCollectionMembership_Valid FM on VRS.ResourceID= FM.ResourceID
                    RIGHT JOIN fn_ListCI_ComplianceState(1033) CI on VRS.ResourceID= CI.ResourceID
                    where CI.ComplianceStateName = 'Compliant' AND CI.CIType_ID = '2'
                    GROUP BY CI.CI_ID) clt ON clt.[CI_ID] = CI.CI_ID
                LEFT JOIN (select CI.CI_ID, COUNT(distinct VRS.Netbios_Name0) as [noncompliancecnt]
                    from v_R_System VRS
                    RIGHT JOIN v_FullCollectionMembership_Valid FM on VRS.ResourceID= FM.ResourceID
                    RIGHT JOIN fn_ListCI_ComplianceState(1033) CI on VRS.ResourceID= CI.ResourceID
                    where CI.ComplianceStateName != 'Compliant' AND CI.CIType_ID = '2'
                    GROUP BY CI.CI_ID) clt2 ON clt2.[CI_ID] = CI.CI_ID
                WHERE CI.CIType_ID = '2' AND clt.[compliancecnt] IS NOT NULL
                GROUP BY CI.CI_ID, CI.CIType_ID, clt.[compliancecnt], clt2.[noncompliancecnt]
                DECLARE @ComplianceCount INT = (SELECT COUNT(*) FROM #ComplianceTable)
                RAISERROR('PROGRESS:18:Processed %d compliance items', 0, 1, @ComplianceCount) WITH NOWAIT
"@

            # Part 5: Create baseline config tables
            if ($Progress) {
                $Progress.UpdateProgress("ProcessSites", "Site $SiteIndex/$TotalSites - Part 5/20: Creating baseline tables",
                    (($SiteIndex - 1) / $TotalSites * 100) + 10)
            }
            $queryParts += @"
                RAISERROR('PROGRESS:19:Creating baseline configuration tables', 0, 1) WITH NOWAIT
                CREATE TABLE #BaselineConfigTable1 (ID INT IDENTITY(1,1), FromCIID int )
                CREATE TABLE #BaselineConfigTable2 (FromCIID int, [BaselineConfigItemRefCount] int)
"@

            # Part 6: Populate BaselineConfigTable1
            if ($Progress) {
                $Progress.UpdateProgress("ProcessSites", "Site $SiteIndex/$TotalSites - Part 6/20: Loading baseline configurations", (($SiteIndex - 1) / $TotalSites * 100) + 12)
            }
            $queryParts += @"
                RAISERROR('PROGRESS:22:Loading baseline configuration items', 0, 1) WITH NOWAIT
                INSERT INTO #BaselineConfigTable1
                SELECT DISTINCT ci.FromCIID FROM (
                    SELECT DISTINCT COUNT(DISTINCT v_DeploymentSummary.CollectionID) as [CollectionCount], COUNT(DISTINCT lci.DisplayName) as [ConfigItemCount], v_LocalizedCIProperties.DisplayName as [BaselineConfig], lci.DisplayName as [ConfigItem], v_CIRelation.FromCIID, v_CIRelation.ToCIID
                    FROM v_ConfigurationItems vcf
                    INNER JOIN v_CITypes ON vcf.CIType_ID = v_CITypes.CIType_ID
                    LEFT JOIN v_LocalizedCIProperties ON vcf.CI_ID = v_LocalizedCIProperties.CI_ID
                    LEFT JOIN v_UpdateInfo ON vcf.CI_ID = v_UpdateInfo.CI_ID
                    LEFT JOIN v_DeploymentSummary ON vcf.[ModelId] = v_DeploymentSummary.[ModelId]
                    LEFT JOIN v_CIRelation ON vcf.CI_ID = v_CIRelation.FromCIID
                    LEFT JOIN v_LocalizedCIProperties lci ON v_CIRelation.ToCIID = lci.CI_ID
                    WHERE vcf.CI_ID IS NOT NULL AND v_LocalizedCIProperties.DisplayName IS NOT NULL AND v_CITypes.TypeName = 'Baseline'
                    GROUP BY v_CITypes.TypeName, v_LocalizedCIProperties.DisplayName,lci.DisplayName,v_CIRelation.FromCIID, v_CIRelation.ToCIID
                    HAVING COUNT(DISTINCT v_DeploymentSummary.CollectionID) > 0) ci
                DECLARE @BaselineCount INT = (SELECT COUNT(*) FROM #BaselineConfigTable1)
                RAISERROR('PROGRESS:25:Loaded %d baseline configurations', 0, 1, @BaselineCount) WITH NOWAIT
"@

            # Part 7-8 COMBINED: Baseline processing loop setup and execution
            if ($Progress) {
                $Progress.UpdateProgress("ProcessSites", "Site $SiteIndex/$TotalSites - Part 7/20: Setting up baseline processing", (($SiteIndex - 1) / $TotalSites * 100) + 14)
            }
            $queryParts += @"
                RAISERROR('PROGRESS:26:Initializing baseline reference processing', 0, 1) WITH NOWAIT

                DECLARE @RowID INT = 1
                DECLARE @RunningTotal BIGINT = 0
                DECLARE @RowCnt BIGINT = 0
                DECLARE @FromCIID int
                DECLARE @BaselineItemConfigRefCount int
                DECLARE @BaselineProgress INT
                DECLARE @ProgressInterval INT

                SELECT @RowCnt = COUNT(0) FROM #BaselineConfigTable1

                -- FIX: Handle case when @RowCnt is 0
                IF @RowCnt = 0
                BEGIN
                    RAISERROR('PROGRESS:45:No baseline configurations to process - skipping', 0, 1) WITH NOWAIT
                END
                ELSE
                BEGIN
                    SET @ProgressInterval = CASE WHEN @RowCnt < 10 THEN 1 ELSE @RowCnt / 10 END

                    RAISERROR('PROGRESS:28:Starting baseline reference processing loop', 0, 1) WITH NOWAIT

                    -- Define CTE once before the loop for efficiency
                    ;WITH BaselineCI AS (
                        SELECT DISTINCT
                            v_CIRelation.FromCIID,
                            v_CIRelation.ToCIID,
                            COUNT(DISTINCT v_DeploymentSummary.CollectionID) as [CollectionCount],
                            COUNT(DISTINCT lci.DisplayName) as [ConfigItemCount],
                            v_LocalizedCIProperties.DisplayName as [BaselineConfig],
                            lci.DisplayName as [ConfigItem]
                        FROM v_ConfigurationItems vcf
                        INNER JOIN v_CITypes ON vcf.CIType_ID = v_CITypes.CIType_ID
                        LEFT JOIN v_LocalizedCIProperties ON vcf.CI_ID = v_LocalizedCIProperties.CI_ID
                        LEFT JOIN v_UpdateInfo ON vcf.CI_ID = v_UpdateInfo.CI_ID
                        LEFT JOIN v_DeploymentSummary ON vcf.[ModelId] = v_DeploymentSummary.[ModelId]
                        LEFT JOIN v_CIRelation ON vcf.CI_ID = v_CIRelation.FromCIID
                        LEFT JOIN v_LocalizedCIProperties lci ON v_CIRelation.ToCIID = lci.CI_ID
                        WHERE vcf.CI_ID IS NOT NULL
                            AND v_LocalizedCIProperties.DisplayName IS NOT NULL
                            AND v_CITypes.TypeName = 'Baseline'
                        GROUP BY v_CITypes.TypeName, v_LocalizedCIProperties.DisplayName,
                            lci.DisplayName, v_CIRelation.FromCIID, v_CIRelation.ToCIID
                        HAVING COUNT(DISTINCT v_DeploymentSummary.CollectionID) > 0
                    )
                    -- Insert the CTE results into a temp table so we can use it in the loop
                    SELECT * INTO #BaselineCITemp FROM BaselineCI

                    -- Now process each baseline item
                    WHILE @RowID <= @RowCnt
                    BEGIN
                        SET @FromCIID = (SELECT FromCIID FROM #BaselineConfigTable1 WHERE [ID] = @RowID)

                        SET @BaselineItemConfigRefCount = (
                            SELECT COUNT(DISTINCT outer_ci.FromCIID)
                            FROM #BaselineCITemp outer_ci
            WHERE outer_ci.ToCIID IN (
                SELECT DISTINCT inner_ci.ToCIID
                FROM #BaselineCITemp inner_ci
                WHERE inner_ci.FromCIID = @FromCIID
            )
        )

        INSERT INTO #BaselineConfigTable2
        SELECT @FromCIID as [FromCIID], @BaselineItemConfigRefCount as [BaselineConfigItemRefCount]

        -- FIX: Progress update with NULL safety
        IF @ProgressInterval > 0 AND @RowID % @ProgressInterval = 0
        BEGIN
            -- Ensure @BaselineProgress is never NULL
            SET @BaselineProgress = 28 + ((@RowID * 17) / @RowCnt)

            -- FIX: Use separate RAISERROR to avoid parameter type issues
            DECLARE @ProgressMsg NVARCHAR(200)
            SET @ProgressMsg = 'PROGRESS:' + CAST(@BaselineProgress AS VARCHAR(10)) + ':Processing baseline reference ' + CAST(@RowID AS VARCHAR(10)) + ' of ' + CAST(@RowCnt AS VARCHAR(10))
            RAISERROR(@ProgressMsg, 0, 1) WITH NOWAIT
        END

        SET @RowID = @RowID + 1
    END

    -- Clean up temp table
    IF OBJECT_ID('tempdb..#BaselineCITemp') IS NOT NULL
        DROP TABLE #BaselineCITemp

    -- Final progress message after loop completes
    RAISERROR('PROGRESS:45:Baseline reference processing complete', 0, 1) WITH NOWAIT
END
"@

        # Part 9: Non-staging collection hierarchy INSERT - Optimized
        if ($Progress) {
            $Progress.UpdateProgress("ProcessSites", "Site $SiteIndex/$TotalSites - Part 9/20: Building collection hierarchy", (($SiteIndex - 1) / $TotalSites * 100) + 18)
        }
        $queryParts += @"
        RAISERROR('PROGRESS:48:Building non-staging collection hierarchy', 0, 1) WITH NOWAIT

        ;WITH CollectionHierarchy AS (
            SELECT ContainerNodeID,
                   CONVERT(nvarchar(max), Name + N'\') AS [Name]
            FROM dbo.Folders
            WHERE ParentContainerNodeID = 0

            UNION ALL

            SELECT fld.ContainerNodeID,
                   cte.Name + fld.Name + N'\'
            FROM dbo.Folders fld
            INNER JOIN CollectionHierarchy cte
                ON fld.ParentContainerNodeID = cte.ContainerNodeID
        )
        INSERT INTO #CollectionHierarchy
        SELECT DISTINCT
            dep.ModelID,
            col.CollectionID,
            col.SiteID AS [CollectionSiteID],
            CAST(0 AS bit) AS [Adaptiva Client Settings Collection Reference],
            CAST(0 AS bit) AS [Adaptiva Policy Collection Reference],
            CAST(0 AS bit) AS [Adaptiva Health Check Collection Reference],
            col.LastMemberChangeTime,
            DATEDIFF(day, col.LastMemberChangeTime, GETDATE()) AS [Days Since Last Member Change],
            col.CollectionName,
            ISNULL(hry.[Name], '/') AS [CollectionPath],
            col.MemberCount,
            CASE WHEN col.CollectionType = 1 THEN 'UserCollection' ELSE 'DeviceCollection' END AS [CollectionType],
            CASE
                WHEN dep.FeatureType = 1 THEN 'Application'
                WHEN dep.FeatureType = 2 THEN 'Package'
                WHEN dep.FeatureType = 6 THEN 'Baseline'
                WHEN dep.FeatureType = 7 THEN 'Task Sequence'
                ELSE 'None'
            END AS [DeploymentType],
            'N/A' AS [baselinetotaldevicecnt],
            'N/A' AS [compliancecnt],
            'N/A' AS [baselinecompliance],
            'N/A' AS [baselinecollectionrefcnt],
            'N/A' AS [baselineconfigitemscount],
            'N/A' AS [baselineconfigitemreference],
            'N/A' AS [baselinecollectionreference],
            ISNULL(app.SoftwareName, 'N/A') AS [ApplicationName],
            CASE WHEN limitingcol.LimitToCollectionID IS NOT NULL THEN 'True' ELSE 'False' END AS [LimitingCollection],
            CASE WHEN refcols.ReferencedCollectionID IS NOT NULL THEN 'True' ELSE 'False' END AS [ReferencedCollection]
        FROM dbo.v_Collections col WITH (NOLOCK)
        LEFT JOIN dbo.v_DeploymentSummary dep WITH (NOLOCK)
            ON col.SiteID = dep.CollectionID
        LEFT JOIN dbo.v_Advertisement adv WITH (NOLOCK)
            ON col.SiteID = adv.CollectionID
        LEFT JOIN dbo.vFolderMembers mem WITH (NOLOCK)
            ON col.SiteID = mem.InstanceKey
        LEFT JOIN CollectionHierarchy hry
            ON mem.ContainerNodeID = hry.ContainerNodeID
        LEFT JOIN (
            SELECT DISTINCT LimitToCollectionID
            FROM v_Collections WITH (NOLOCK)
            WHERE LimitToCollectionID IS NOT NULL
        ) limitingcol ON limitingcol.LimitToCollectionID = col.SiteID
        LEFT JOIN (
            SELECT DISTINCT ReferencedCollectionID
            FROM Collection_Rules WITH (NOLOCK)
            WHERE ReferencedCollectionID IS NOT NULL
        ) refcols ON refcols.ReferencedCollectionID = col.SiteID
        LEFT JOIN (
            SELECT SoftwareName, CollectionName
            FROM fn_DeploymentSummary(1033)
            WHERE FeatureType = 1
        ) app ON app.CollectionName = col.CollectionName
        WHERE adv.AdvertisementID IS NULL
            AND col.PowerConfigsCount = '0'
            AND col.SiteID NOT IN (
                SELECT CollectionID
                FROM vSMS_ClientSettingsAssignments WITH (NOLOCK)
            )
            AND mem.ObjectType IN ('5000', '5001')
            AND (hry.[Name] NOT LIKE '%Staging%' OR hry.[Name] IS NULL)
            AND (hry.[Name] NOT LIKE '%Retired%' OR hry.[Name] IS NULL)
        OPTION (MAXDOP 1, RECOMPILE, OPTIMIZE FOR UNKNOWN);

        DECLARE @NonStagingCount INT = (SELECT COUNT(*) FROM #CollectionHierarchy)
        RAISERROR('PROGRESS:55:Loaded %d non-staging collections', 0, 1, @NonStagingCount) WITH NOWAIT
"@

        # Part 10: Staging CTE start - Optimized
        if ($Progress) {
            $Progress.UpdateProgress("ProcessSites", "Site $SiteIndex/$TotalSites - Part 10/20: Building staging hierarchy", (($SiteIndex - 1) / $TotalSites * 100) + 20)
        }
        $queryParts += @"
        RAISERROR('PROGRESS:58:Building staging collection hierarchy', 0, 1) WITH NOWAIT

        ;WITH CollectionHierarchyStaging AS (
            SELECT ContainerNodeID,
                   CONVERT(nvarchar(max), Name + N'\') AS [Name]
            FROM dbo.Folders
            WHERE ParentContainerNodeID = 0

            UNION ALL

            SELECT fld.ContainerNodeID,
                   cte.Name + fld.Name + N'\'
            FROM dbo.Folders fld
            INNER JOIN CollectionHierarchyStaging cte
                ON fld.ParentContainerNodeID = cte.ContainerNodeID
        )
        INSERT INTO #CollectionHierarchyStaging
        SELECT DISTINCT
            dep.ModelID,
            col.CollectionID,
            col.SiteID AS [CollectionSiteID],
            CAST(0 AS bit) AS [Adaptiva Client Settings Collection Reference],
            CAST(0 AS bit) AS [Adaptiva Policy Collection Reference],
            CAST(0 AS bit) AS [Adaptiva Health Check Collection Reference],
            col.LastMemberChangeTime,
            DATEDIFF(day, col.LastMemberChangeTime, GETDATE()) AS [Days Since Last Member Change],
            col.CollectionName,
            ISNULL(hry.[Name], '/') AS [CollectionPath],
            col.MemberCount,
            CASE WHEN col.CollectionType = 1 THEN 'UserCollection' ELSE 'DeviceCollection' END AS [CollectionType],
            CASE
                WHEN dep.FeatureType = 1 THEN 'Application'
                WHEN dep.FeatureType = 2 THEN 'Package'
                WHEN dep.FeatureType = 6 THEN 'Baseline'
                WHEN dep.FeatureType = 7 THEN 'Task Sequence'
                ELSE 'None'
            END AS [DeploymentType],
            'N/A' AS [baselinetotaldevicecnt],
            'N/A' AS [compliancecnt],
            'N/A' AS [baselinecompliance],
            'N/A' AS [baselinecollectionrefcnt],
            'N/A' AS [baselineconfigitemscount],
            'N/A' AS [baselineconfigitemreference],
            'N/A' AS [baselinecollectionreference],
            ISNULL(app.SoftwareName, 'N/A') AS [ApplicationName],
            CASE WHEN limitingcol.LimitToCollectionID IS NOT NULL THEN 'True' ELSE 'False' END AS [LimitingCollection],
            CASE WHEN refcols.ReferencedCollectionID IS NOT NULL THEN 'True' ELSE 'False' END AS [ReferencedCollection]
        FROM dbo.v_Collections col WITH (NOLOCK)
        LEFT JOIN dbo.v_DeploymentSummary dep WITH (NOLOCK)
            ON col.SiteID = dep.CollectionID
        LEFT JOIN dbo.v_Advertisement adv WITH (NOLOCK)
            ON col.SiteID = adv.CollectionID
        LEFT JOIN dbo.vFolderMembers mem WITH (NOLOCK)
            ON col.SiteID = mem.InstanceKey
        LEFT JOIN CollectionHierarchyStaging hry
            ON mem.ContainerNodeID = hry.ContainerNodeID
        LEFT JOIN (
            SELECT DISTINCT LimitToCollectionID
            FROM v_Collections WITH (NOLOCK)
            WHERE LimitToCollectionID IS NOT NULL
        ) limitingcol ON limitingcol.LimitToCollectionID = col.SiteID
        LEFT JOIN (
            SELECT DISTINCT ReferencedCollectionID
            FROM Collection_Rules WITH (NOLOCK)
            WHERE ReferencedCollectionID IS NOT NULL
        ) refcols ON refcols.ReferencedCollectionID = col.SiteID
        LEFT JOIN (
            SELECT SoftwareName, CollectionName
            FROM fn_DeploymentSummary(1033)
            WHERE FeatureType = 1
        ) app ON app.CollectionName = col.CollectionName
        WHERE adv.AdvertisementID IS NULL
            AND col.PowerConfigsCount = '0'
            AND col.SiteID NOT IN (
                SELECT CollectionID
                FROM vSMS_ClientSettingsAssignments WITH (NOLOCK)
            )
            AND mem.ObjectType IN ('5000', '5001')
            AND hry.[Name] LIKE '%Staging%'
            AND hry.[Name] NOT LIKE '%Retired%'
        OPTION (MAXDOP 1, RECOMPILE, OPTIMIZE FOR UNKNOWN);

        DECLARE @StagingCount INT = (SELECT COUNT(*) FROM #CollectionHierarchyStaging)
        RAISERROR('PROGRESS:65:Loaded %d staging collections', 0, 1, @StagingCount) WITH NOWAIT
"@

        # Part 11: Non-staging final SELECT - Start
        if ($Progress) {
            $Progress.UpdateProgress("ProcessSites", "Site $SiteIndex/$TotalSites - Part 11/20: Selecting non-staging collections", (($SiteIndex - 1) / $TotalSites * 100) + 22)
        }
        $queryParts += @"
        RAISERROR('PROGRESS:66:Preparing non-staging collection output', 0, 1) WITH NOWAIT
        SELECT DISTINCT ch.CollectionID, ch.CollectionSiteID, ch.[Adaptiva Client Settings Collection Reference],
            ch.[Adaptiva Policy Collection Reference], ch.[Adaptiva Health Check Collection Reference],
            ch.LastMemberChangeTime, ch.[Days Since Last Member Change], ch.CollectionName, ch.CollectionPath,
            ch.MemberCount, ch.CollectionType, ch.DeploymentType,
            ch.baselinetotaldevicecnt, ch.compliancecnt, ch.baselinecompliance, ch.baselinecollectionrefcnt,
            ch.baselineconfigitemscount,
            ch.baselineconfigitemreference, ch.baselinecollectionreference, ch.ApplicationName, ch.LimitingCollection,
            ch.ReferencedCollection, CASE WHEN dep.[DeploymentCount] IS NULL THEN 0 ELSE dep.[DeploymentCount] END AS [DeploymentCount],
            CASE WHEN appobj.ObjectPath IS NULL THEN CASE WHEN objpath.folderpath IS NULL THEN 'N/A' ELSE objpath.folderpath END ELSE appobj.ObjectPath END as [FolderPath],
            CASE WHEN appobj.[Package ID] IS NULL THEN 'N/A' ELSE appobj.[Package ID] END as [Package ID],
            CASE WHEN CAST(spk.SourceSize as varchar) IS NULL THEN 'N/A' ELSE CAST(spk.SourceSize as varchar) END as [PackageSizekB],
            CASE WHEN appobj.[Package Type] IS NULL THEN 'N/A' ELSE appobj.[Package Type] END as [Package Type],
            CASE WHEN tsref.[TS Name] IS NOT NULL THEN tsref.[TS Name] ELSE 'None Found' END as [TS Name],
            CASE WHEN appdep.[Application] IS NULL THEN 'False' ELSE 'True' END as [App TS References],
            CASE WHEN appdeps.[ChildAppDisplayName] IS NULL THEN 'False' ELSE 'True' END as [App References],
            CASE WHEN Applications.[AdminComments] IS NULL THEN 'N/A' ELSE Applications.[AdminComments] END as [App_Comments]
        FROM #CollectionHierarchy ch
        LEFT JOIN [dbo].[v_Applications] Applications on Applications.[DisplayName] = ch.ApplicationName
        LEFT JOIN #appObjects appobj ON appobj.[Package Name] = ch.ApplicationName
"@

        # Part 12: Non-staging final SELECT - Configuration Items joins
        if ($Progress) {
            $Progress.UpdateProgress("ProcessSites", "Site $SiteIndex/$TotalSites - Part 12/20: Joining configuration items", (($SiteIndex - 1) / $TotalSites * 100) + 24)
        }
        $queryParts += @"
        LEFT JOIN (SELECT DISTINCT vcf.CI_ID, vcf.ModelID, v_CITypes.TypeName, v_CITypes.CIType_ID,
                   v_LocalizedCIProperties.DisplayName
            FROM v_ConfigurationItems vcf
            INNER JOIN v_CITypes ON vcf.CIType_ID = v_CITypes.CIType_ID
            LEFT JOIN v_LocalizedCIProperties ON vcf.CI_ID = v_LocalizedCIProperties.CI_ID
            LEFT JOIN v_UpdateInfo ON vcf.CI_ID = v_UpdateInfo.CI_ID ) basecic ON basecic.ModelID = ch.ModelID
        LEFT JOIN (SELECT DISTINCT [ci_id], [displayname] FROM v_LocalizedCIProperties) ciid ON (ciid.displayname = ch.ApplicationName) OR (ciid.displayname = basecic.[DisplayName])
        LEFT JOIN (SELECT f.FolderPath, cfi.ci_id from v_ConfigurationItems cfi
            LEFT JOIN [dbo].[FolderMembers] fm WITH (NOLOCK) ON fm.InstanceKey = cfi.CI_UniqueID
            LEFT JOIN [dbo].[Folders] f WITH (NOLOCK) ON fm.ContainerNodeID = f.ContainerNodeID) objpath ON objpath.ci_id = ciid.ci_id
"@

        # Part 13: Non-staging final SELECT - Package and TS joins
        if ($Progress) {
            $Progress.UpdateProgress("ProcessSites", "Site $SiteIndex/$TotalSites - Part 13/20: Joining packages and task sequences", (($SiteIndex - 1) / $TotalSites * 100) + 26)
        }
        $queryParts += @"
        LEFT JOIN (SELECT DISTINCT p.Name, p.[Description], n.SourceCompressedSize, n.SourceSize, dp.LastRefreshTime,
                   p.Manufacturer, p.[Version],
                   p.[Language], p.SourceSite, p.PackageID, CASE WHEN dp.IsPeerDP=1 THEN '*' ELSE '' END as BranchDP,
                   psd.InstallStatus
            FROM v_Package p
            INNER JOIN v_DistributionPoint dp ON p.PackageID = dp.PackageID
            LEFT JOIN v_PackageStatusRootSummarizer n ON p.PackageID = n.PackageID
            LEFT JOIN v_PackageStatusDistPointsSumm psd ON dp.ServerNALPath=psd.ServerNALPath AND dp.PackageID=psd.PackageID
            LEFT JOIN v_PackageStatus ps
                ON dp.ServerNALPath=ps.PkgServer AND dp.PackageID=ps.PackageID) spk ON spk.PackageID = appobj.[Package ID]
        LEFT JOIN (SELECT DISTINCT pkg.PackageID as [AppID], pkg.[Name] as [AppName],
                   CASE pkg.PackageType
                       WHEN 0 THEN 'Software Distribution Package'
                       WHEN 3 THEN 'Driver Package'
                       WHEN 4 THEN 'Task Sequence Package'
                       WHEN 5 THEN 'Software Update Package'
                       WHEN 6 THEN 'Device Settings Package'
                       WHEN 7 THEN 'Virtual Package'
                       WHEN 257 THEN 'Image Package'
                       WHEN 258 THEN 'Boot Image Package'
                       WHEN 259 THEN 'OS Install Package'
                   END AS [AppType], tsp.PackageID as [TS AppID], tsp.[Name] as [TS Name]
            FROM v_Package pkg
            LEFT JOIN v_Advertisement adv on pkg.PackageID=adv.PackageID
            LEFT JOIN v_TaskSequencePackageReferences tsr on pkg.PackageID=tsr.RefPackageID
            LEFT JOIN v_TaskSequencePackage tsp on tsr.PackageID=tsp.PackageID
            WHERE adv.AdvertisementID is null
            AND pkg.PackageType <> 8) tsref ON ch.ApplicationName = tsref.[AppName]
"@

        # Part 14: Non-staging final SELECT - Application dependencies
        if ($Progress) {
            $Progress.UpdateProgress("ProcessSites", "Site $SiteIndex/$TotalSites - Part 14/20: Joining application dependencies", (($SiteIndex - 1) / $TotalSites * 100) + 28)
        }
        $queryParts += @"
        LEFT JOIN (select DISTINCT app.DisplayName [Application],TS.[Name] [TS Name],TS.PackageID [TS_ID],TS.Description [TSDescription] From dbo.fn_ListLatestApplicationCIs(1033) app
                   inner join v_TaskSequenceAppReferencesInfo TSApp on app.ModelName=TSApp.RefAppModelName
                   inner join v_TaskSequencePackage TS on TS.PackageID=TSApp.PackageID) appdep ON appdep.[Application] = ch.ApplicationName
        LEFT JOIN (select DISTINCT parentapp.displayname [Application Name],
                   parentapp.CreatedBy [Created By],
                   ParentApp_F.Name AS [ParentAppFolder] ,
                   ParentApp_DT.Technology AS [ParentAppDeploymentType],
                   case when parentapp.IsEnabled='1' then 'Yes' else 'No' end as IsEnabled,
                   case when parentapp.IsDeployed='1' then 'Yes' else 'No' end as IsDeployed,
                   parentapp.NumberOfDeploymentTypes [No of DT],
                   parentapp.NumberOfDeployments [No of Deployments],
                   parentapp.NumberOfDevicesWithApp,
                   parentapp.NumberOfDevicesWithFailure,
                   ChildApp.DisplayName as [ChildAppDisplayName]
            From dbo.fn_ListApplicationCIs(1033) ParentApp
            Left Join dbo.vFolderMembers ParentApp_FM on ParentApp_FM.InstanceKey = ParentApp.ModelName
                   Left Join dbo.vSMS_Folders ParentApp_F on ParentApp_F.ContainerNodeID = ParentApp_FM.ContainerNodeID
                   Left Join fn_ListDeploymentTypeCIs(1033) ParentApp_DT ON ParentApp_DT.AppModelName = ParentApp.ModelName
                   Left Join vSMS_AppRelation_Flat R on R.FromApplicationCIID = ParentApp.CI_ID
                   Left Join fn_ListApplicationCIs_List(1033) ChildApp on ChildApp.CI_ID = R.ToApplicationCIID
                        And ChildApp.IsLatest = 1
                   Left Join dbo.vFolderMembers ChildApp_FM on ChildApp_FM.InstanceKey =ChildApp.ModelName
                   Left Join dbo.vSMS_Folders ChildApp_F on ChildApp_F.ContainerNodeID = ChildApp_FM.ContainerNodeID
                   Left Join fn_ListDeploymentTypeCIs(1033) ChildApp_DT ON ChildApp_DT.AppModelName = ChildApp.ModelName
            where parentapp.IsLatest='1' AND ChildApp.DisplayName IS NOT NULL
            group by
                   parentapp.DisplayName,parentapp.CreatedBy,parentapp.NumberOfDeploymentTypes,parentapp.NumberOfDeployments,
                   parentapp.NumberOfDevicesWithApp,
                   parentapp.NumberOfDevicesWithFailure,
                   ChildApp.DisplayName,
                   parentapp.IsEnabled,
                   parentapp.IsDeployed,
                   ParentApp_F.Name,
                   ParentApp_DT.Technology,
                   ParentApp.IsSuperseded
        ) appdeps ON appdeps.[ChildAppDisplayName] = ch.ApplicationName
"@

        # Part 15: Non-staging final SELECT - Deployment count
        if ($Progress) {
            $Progress.UpdateProgress("ProcessSites", "Site $SiteIndex/$TotalSites - Part 15/20: Calculating deployment counts", (($SiteIndex - 1) / $TotalSites * 100) + 30)
        }
        $queryParts += @"
        LEFT JOIN (SELECT DISTINCT APP.DisplayName [App Name], COUNT(CIA.CollectionID) as [DeploymentCount]
            FROM fn_ListApplicationCIs(1033) APP
            JOIN fn_ListDeploymentTypeCIs(1033) DT ON DT.AppModelName = APP.ModelName AND DT.IsLatest = 1
            LEFT JOIN v_CIAssignmentToCI CIACI ON CIACI.CI_ID = APP.CI_ID
            LEFT JOIN v_CIAssignment CIA ON CIACI.AssignmentID = CIA.AssignmentID
            LEFT JOIN v_Collection C ON C.CollectionID = CIA.CollectionID
            WHERE APP.IsLatest = 1
            GROUP BY APP.DisplayName) dep ON dep.[App Name] = ch.ApplicationName
"@

        # Part 16: Staging final SELECT - Start
        if ($Progress) {
            $Progress.UpdateProgress("ProcessSites", "Site $SiteIndex/$TotalSites - Part 16/20: Selecting staging collections", (($SiteIndex - 1) / $TotalSites * 100) + 32)
        }
        $queryParts += @"
        RAISERROR('PROGRESS:81:Preparing staging collection output', 0, 1) WITH NOWAIT
        SELECT DISTINCT ch.CollectionID, ch.CollectionSiteID, ch.[Adaptiva Client Settings Collection Reference],
            ch.[Adaptiva Policy Collection Reference], ch.[Adaptiva Health Check Collection Reference],
            ch.LastMemberChangeTime, ch.[Days Since Last Member Change], ch.CollectionName, ch.CollectionPath,
            ch.MemberCount, ch.CollectionType, ch.DeploymentType,
            ch.baselinetotaldevicecnt, ch.compliancecnt, ch.baselinecompliance, ch.baselinecollectionrefcnt,
            ch.baselineconfigitemscount,
            ch.baselineconfigitemreference, ch.baselinecollectionreference, ch.ApplicationName, ch.LimitingCollection,
            ch.ReferencedCollection, CASE WHEN dep.[DeploymentCount] IS NULL THEN 0 ELSE dep.[DeploymentCount] END AS [DeploymentCount],
            CASE WHEN appobj.ObjectPath IS NULL THEN CASE WHEN objpath.folderpath IS NULL THEN 'N/A' ELSE objpath.folderpath END ELSE appobj.ObjectPath END as [FolderPath],
            CASE WHEN appobj.[Package ID] IS NULL THEN 'N/A' ELSE appobj.[Package ID] END as [Package ID],
            CASE WHEN CAST(spk.SourceSize as varchar) IS NULL THEN 'N/A' ELSE CAST(spk.SourceSize as varchar) END as [PackageSizekB],
            CASE WHEN appobj.[Package Type] IS NULL THEN 'N/A' ELSE appobj.[Package Type] END as [Package Type],
            CASE WHEN tsref.[TS Name] IS NOT NULL THEN tsref.[TS Name] ELSE 'None Found' END as [TS Name],
            CASE WHEN appdep.[Application] IS NULL THEN 'False' ELSE 'True' END as [App TS References],
            CASE WHEN appdeps.[ChildAppDisplayName] IS NULL THEN 'False' ELSE 'True' END as [App References],
            CASE WHEN Applications.[AdminComments] IS NULL THEN 'N/A' ELSE Applications.[AdminComments] END as [App_Comments]
        FROM #CollectionHierarchyStaging ch
        LEFT JOIN [dbo].[v_Applications] Applications on Applications.[DisplayName] = ch.ApplicationName
        LEFT JOIN #appObjects appobj ON appobj.[Package Name] = ch.ApplicationName
"@

        # Part 17: Staging final SELECT - Configuration Items joins
        if ($Progress) {
            $Progress.UpdateProgress("ProcessSites", "Site $SiteIndex/$TotalSites - Part 17/20: Joining staging configuration items", (($SiteIndex - 1) / $TotalSites * 100) + 34)
        }
        $queryParts += @"
        LEFT JOIN (SELECT DISTINCT vcf.CI_ID, vcf.ModelID, v_CITypes.TypeName, v_CITypes.CIType_ID,
                   v_LocalizedCIProperties.DisplayName
            FROM v_ConfigurationItems vcf
            INNER JOIN v_CITypes ON vcf.CIType_ID = v_CITypes.CIType_ID
            LEFT JOIN v_LocalizedCIProperties ON vcf.CI_ID = v_LocalizedCIProperties.CI_ID
            LEFT JOIN v_UpdateInfo ON vcf.CI_ID = v_UpdateInfo.CI_ID ) basecic ON basecic.ModelID = ch.ModelID
        LEFT JOIN (SELECT DISTINCT [ci_id], [displayname] FROM v_LocalizedCIProperties) ciid ON (ciid.displayname = ch.ApplicationName) OR (ciid.displayname = basecic.[DisplayName])
        LEFT JOIN (SELECT f.FolderPath, cfi.ci_id from v_ConfigurationItems cfi
            LEFT JOIN [dbo].[FolderMembers] fm WITH (NOLOCK) ON fm.InstanceKey = cfi.CI_UniqueID
            LEFT JOIN [dbo].[Folders] f WITH (NOLOCK) ON fm.ContainerNodeID = f.ContainerNodeID) objpath ON objpath.ci_id = ciid.ci_id
"@

        # Part 18: Staging final SELECT - Package and TS joins
        if ($Progress) {
            $Progress.UpdateProgress("ProcessSites", "Site $SiteIndex/$TotalSites - Part 18/20: Joining staging packages and task sequences", (($SiteIndex - 1) / $TotalSites * 100) + 36)
        }
        $queryParts += @"
        LEFT JOIN (SELECT DISTINCT p.Name, p.[Description], n.SourceCompressedSize, n.SourceSize, dp.LastRefreshTime,
                   p.Manufacturer, p.[Version],
                   p.[Language], p.SourceSite, p.PackageID, CASE WHEN dp.IsPeerDP=1 THEN '*' ELSE '' END as BranchDP,
                   psd.InstallStatus
            FROM v_Package p
            INNER JOIN v_DistributionPoint dp ON p.PackageID = dp.PackageID
            LEFT JOIN v_PackageStatusRootSummarizer n ON p.PackageID = n.PackageID
            LEFT JOIN v_PackageStatusDistPointsSumm psd ON dp.ServerNALPath=psd.ServerNALPath AND dp.PackageID=psd.PackageID
            LEFT JOIN v_PackageStatus ps
                ON dp.ServerNALPath=ps.PkgServer AND dp.PackageID=ps.PackageID) spk ON spk.PackageID = appobj.[Package ID]
        LEFT JOIN (SELECT DISTINCT pkg.PackageID as [AppID], pkg.[Name] as [AppName],
                   CASE pkg.PackageType
                       WHEN 0 THEN 'Software Distribution Package'
                       WHEN 3 THEN 'Driver Package'
                       WHEN 4 THEN 'Task Sequence Package'
                       WHEN 5 THEN 'Software Update Package'
                       WHEN 6 THEN 'Device Settings Package'
                       WHEN 7 THEN 'Virtual Package'
                       WHEN 257 THEN 'Image Package'
                       WHEN 258 THEN 'Boot Image Package'
                       WHEN 259 THEN 'OS Install Package'
                   END AS [AppType], tsp.PackageID as [TS AppID], tsp.[Name] as [TS Name]
            FROM v_Package pkg
            LEFT JOIN v_Advertisement adv on pkg.PackageID=adv.PackageID
            LEFT JOIN v_TaskSequencePackageReferences tsr on pkg.PackageID=tsr.RefPackageID
            LEFT JOIN v_TaskSequencePackage tsp on tsr.PackageID=tsp.PackageID
            WHERE adv.AdvertisementID is null
            AND pkg.PackageType <> 8) tsref ON ch.ApplicationName = tsref.[AppName]
"@

        # Part 19: Staging final SELECT - Application dependencies
        if ($Progress) {
            $Progress.UpdateProgress("ProcessSites", "Site $SiteIndex/$TotalSites - Part 19/20: Joining staging application dependencies", (($SiteIndex - 1) / $TotalSites * 100) + 38)
        }
        $queryParts += @"
        LEFT JOIN (select DISTINCT app.DisplayName [Application],TS.[Name] [TS Name],TS.PackageID [TS_ID],TS.Description [TSDescription] From dbo.fn_ListLatestApplicationCIs(1033) app
                   inner join v_TaskSequenceAppReferencesInfo TSApp on app.ModelName=TSApp.RefAppModelName
                   inner join v_TaskSequencePackage TS on TS.PackageID=TSApp.PackageID) appdep ON appdep.[Application] = ch.ApplicationName
        LEFT JOIN (select DISTINCT parentapp.displayname [Application Name],
                   parentapp.CreatedBy [Created By],
                   ParentApp_F.Name AS [ParentAppFolder] ,
                   ParentApp_DT.Technology AS [ParentAppDeploymentType],
                   case when parentapp.IsEnabled='1' then 'Yes' else 'No' end as IsEnabled,
                   case when parentapp.IsDeployed='1' then 'Yes' else 'No' end as IsDeployed,
                   parentapp.NumberOfDeploymentTypes [No of DT],
                   parentapp.NumberOfDeployments [No of Deployments],
                   parentapp.NumberOfDevicesWithApp,
                   parentapp.NumberOfDevicesWithFailure,
                   ChildApp.DisplayName as [ChildAppDisplayName]
            From dbo.fn_ListApplicationCIs(1033) ParentApp
            Left Join dbo.vFolderMembers ParentApp_FM on ParentApp_FM.InstanceKey = ParentApp.ModelName
                   Left Join dbo.vSMS_Folders ParentApp_F on ParentApp_F.ContainerNodeID = ParentApp_FM.ContainerNodeID
                   Left Join fn_ListDeploymentTypeCIs(1033) ParentApp_DT ON ParentApp_DT.AppModelName = ParentApp.ModelName
                   Left Join vSMS_AppRelation_Flat R on R.FromApplicationCIID = ParentApp.CI_ID
                   Left Join fn_ListApplicationCIs_List(1033) ChildApp on ChildApp.CI_ID = R.ToApplicationCIID
                        And ChildApp.IsLatest = 1
                   Left Join dbo.vFolderMembers ChildApp_FM on ChildApp_FM.InstanceKey =ChildApp.ModelName
                   Left Join dbo.vSMS_Folders ChildApp_F on ChildApp_F.ContainerNodeID = ChildApp_FM.ContainerNodeID
                   Left Join fn_ListDeploymentTypeCIs(1033) ChildApp_DT ON ChildApp_DT.AppModelName = ChildApp.ModelName
            where parentapp.IsLatest='1' AND ChildApp.DisplayName IS NOT NULL
            group by
                   parentapp.DisplayName,parentapp.CreatedBy,parentapp.NumberOfDeploymentTypes,parentapp.NumberOfDeployments,
                   parentapp.NumberOfDevicesWithApp,
                   parentapp.NumberOfDevicesWithFailure,
                   ChildApp.DisplayName,
                   parentapp.IsEnabled,
                   parentapp.IsDeployed,
                   ParentApp_F.Name,
                   ParentApp_DT.Technology,
                   ParentApp.IsSuperseded
        ) appdeps ON appdeps.[ChildAppDisplayName] = ch.ApplicationName
"@

        # Part 20: Staging final SELECT - Deployment count and cleanup
        if ($Progress) {
            $Progress.UpdateProgress("ProcessSites", "Site $SiteIndex/$TotalSites - Part 20/20: Finalizing and cleanup", (($SiteIndex - 1) / $TotalSites * 100) + 40)
        }
        $queryParts += @"
        LEFT JOIN (SELECT DISTINCT APP.DisplayName [App Name], COUNT(CIA.CollectionID) as [DeploymentCount]
            FROM fn_ListApplicationCIs(1033) APP
            JOIN fn_ListDeploymentTypeCIs(1033) DT ON DT.AppModelName = APP.ModelName AND DT.IsLatest = 1
            LEFT JOIN v_CIAssignmentToCI CIACI ON CIACI.CI_ID = APP.CI_ID
            LEFT JOIN v_CIAssignment CIA ON CIACI.AssignmentID = CIA.AssignmentID
            LEFT JOIN v_Collection C ON C.CollectionID = CIA.CollectionID
            WHERE APP.IsLatest = 1
            GROUP BY APP.DisplayName) dep ON dep.[App Name] = ch.ApplicationName;

        RAISERROR('PROGRESS:97:Cleaning up temporary tables', 0, 1) WITH NOWAIT

        DROP TABLE #BaselineConfigTable1
        DROP TABLE #BaselineConfigTable2
        DROP TABLE #ComplianceTable
        DROP TABLE #CollectionHierarchy
        DROP TABLE #CollectionHierarchyStaging
        DROP TABLE #appObjects

        RAISERROR('PROGRESS:100:Collection detection query complete', 0, 1) WITH NOWAIT
"@

        # Combine all query parts into single query
        $Logger.Info("Combining query parts into final query", "CollectionDetectionQuery")
        $fullQuery = $queryParts -join "`n`n"

        # Add SET NOCOUNT ON at the beginning if not already present
        if ($fullQuery -notmatch "SET NOCOUNT ON") {
            $fullQuery = "SET NOCOUNT ON`n" + $fullQuery
        }

        # Output query to text file for SSMS validation
        $queryOutputPath = Join-Path $Config.LogDir "CollectionDetectionQuery_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').sql"
        try {
            $fullQuery | Out-File -FilePath $queryOutputPath -Encoding UTF8 -Force
            $Logger.Info("Query exported to: $queryOutputPath", "CollectionDetectionQuery")
            Write-Host ""  # Add blank line
            Write-Host "SQL Query saved to: $queryOutputPath" -ForegroundColor Cyan
        }
        catch {
            $Logger.Warning("Failed to export query to file: $_", "CollectionDetectionQuery")
        }

        # Log query statistics
        $queryLength = $fullQuery.Length
        $queryLineCount = ($fullQuery -split "`n").Count
        $Logger.Info("Query assembled - Length: $queryLength chars, Lines: $queryLineCount", "CollectionDetectionQuery")

        # Execute the combined query
        if ($Progress) {
            $Progress.UpdateProgress("ProcessSites",
                "Site $SiteIndex/$TotalSites - Executing collection detection query",
                (($SiteIndex - 1) / $TotalSites * 100) + 50)
        }

        $Logger.Info("Executing combined collection detection query", "CollectionDetectionQuery")

        $dataset = New-Object System.Data.DataSet
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($fullQuery, $SCCMConnection.Connection)
        $adapter.SelectCommand.CommandTimeout = 21600  # 6 hour timeout

        $ResourceManager.Register($adapter, "CollectionDetectionAdapter")

        $queryStartTime = Get-Date
        [void]$adapter.Fill($dataset)
        $queryEndTime = Get-Date
        $queryDuration = ($queryEndTime - $queryStartTime).TotalSeconds

        $Logger.Info("Collection detection query completed in $([math]::Round($queryDuration, 2)) seconds", "CollectionDetectionQuery")

        # Log all progress messages received
        if ($progressMessages.Count -gt 0) {
            $Logger.Info("Query progress messages received: $($progressMessages.Count)", "CollectionDetectionQuery")
            foreach ($msg in $progressMessages) {
                $Logger.Debug($msg, "CollectionDetectionQuery")
            }
        }

        # Validate dataset
        if ($dataset.Tables.Count -eq 0) {
            throw "Query returned no result sets"
        }

        # Merge results from both queries (Non-Staging and Staging)
        $dt = $dataset.Tables[0]

        if ($dataset.Tables.Count -gt 1) {
            $Logger.Info("Merging staging collection results", "CollectionDetectionQuery")
            $dt.Merge($dataset.Tables[1])
        }

        $Logger.Info("Total collections retrieved: $($dt.Rows.Count)", "CollectionDetectionQuery")

        # Update final progress
        if ($Progress) {
            $Progress.UpdateProgress("ProcessSites",
                "Site $SiteIndex/$TotalSites - Collection detection complete",
                (($SiteIndex - 1) / $TotalSites * 100) + 60)
        }

        return @{
            Success = $true
            Data = $dt
        }
    }
    finally
    {
        # NEW: Always remove event handler in finally block
        try {
            $SCCMConnection.Connection.remove_InfoMessage($infoMessageHandler)
            $Logger.Debug("InfoMessage event handler removed", "CollectionDetectionQuery")
        }
        catch {
            $Logger.Warning("Failed to remove InfoMessage event handler: $_", "CollectionDetectionQuery")
        }
    }
}
catch
{
    $ErrorTracker.RecordError("CollectionDetectionQuery",
        "Failed to execute Collection Detection query",
        $false,
        $_.Exception)
    $Logger.Error("Collection Detection query failed: $_", "CollectionDetectionQuery")

    return @{
        Success = $false
        Data = $null
    }
}
}
#endregion

#region Retired Folder ScannerPoint
function Invoke-RetiredFolderScan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [DatabaseConnection]$SCCMConnection,

        [Parameter(Mandatory = $true)]
        [Logger]$Logger,

        [Parameter(Mandatory = $true)]
        [ErrorTracker]$ErrorTracker,

        [Parameter(Mandatory = $true)]
        [ResourceManager]$ResourceManager,

        [Parameter(Mandatory = $false)]
        [ProgressTracker]$Progress,

        [Parameter(Mandatory = $false)]
        [int]$SiteIndex = 1,

        [Parameter(Mandatory = $false)]
        [int]$TotalSites = 1
    )

    try {
        $Logger.Info("Starting retired folder scan", "RetiredFolderScan")

        # Query to find all objects in folders containing "retired" in the path
        $retiredQuery = @"
SET NOCOUNT ON

-- CTE for folder hierarchy
;WITH FolderHierarchy AS (
    SELECT ContainerNodeID,
           CONVERT(nvarchar(max), Name + N'\') AS [Path],
           Name
    FROM dbo.Folders
    WHERE ParentContainerNodeID = 0

    UNION ALL

    SELECT fld.ContainerNodeID,
           cte.Path + fld.Name + N'\',
           fld.Name
    FROM dbo.Folders fld
    INNER JOIN FolderHierarchy cte
        ON fld.ParentContainerNodeID = cte.ContainerNodeID
)
-- Get User Collections in Retired folders
SELECT DISTINCT
    'UserCollection' AS [ObjectType],
    CAST(col.CollectionID AS nvarchar(50)) AS [CollectionID],
    CAST(col.CollectionID AS nvarchar(50)) AS [CollectionNumericID],
    col.SiteID AS [CollectionSiteID],
    col.CollectionName AS [Name],
    ISNULL(hry.[Path], '/') AS [FolderPath],
    col.MemberCount,
    col.LastMemberChangeTime,
    DATEDIFF(day, col.LastMemberChangeTime, GETDATE()) AS [DaysSinceLastChange],
    CASE WHEN col.LimitToCollectionID IS NOT NULL THEN 'True' ELSE 'False' END AS [HasLimitingCollection],
    col.LimitToCollectionID AS [LimitingCollectionID],
    col.LimitToCollectionName AS [LimitingCollectionName],
    NULL AS [PackageID],
    NULL AS [PackageName],
    NULL AS [PackageType],
    NULL AS [ReferencePackageID],
    NULL AS [ReferencedPackageName],
    NULL AS [CollectionName],
    NULL AS [CollectionType],
    NULL AS [CollectionObjectPath]
FROM dbo.v_Collections col WITH (NOLOCK)
LEFT JOIN dbo.vFolderMembers mem WITH (NOLOCK)
    ON col.SiteID = mem.InstanceKey
LEFT JOIN FolderHierarchy hry
    ON mem.ContainerNodeID = hry.ContainerNodeID
LEFT JOIN dbo.v_Advertisement adv WITH (NOLOCK)
    ON col.SiteID = adv.CollectionID
WHERE col.CollectionType = 1
    AND mem.ObjectType = '5001'
    AND hry.[Path] LIKE '%Retired%'
    AND adv.AdvertisementID IS NULL

UNION ALL

-- Get Device Collections in Retired folders
SELECT DISTINCT
    'DeviceCollection' AS [ObjectType],
    CAST(col.CollectionID AS nvarchar(50)) AS [CollectionID],
    CAST(col.CollectionID AS nvarchar(50)) AS [CollectionNumericID],
    col.SiteID AS [CollectionSiteID],
    col.CollectionName AS [Name],
    ISNULL(hry.[Path], '/') AS [FolderPath],
    col.MemberCount,
    col.LastMemberChangeTime,
    DATEDIFF(day, col.LastMemberChangeTime, GETDATE()) AS [DaysSinceLastChange],
    CASE WHEN col.LimitToCollectionID IS NOT NULL THEN 'True' ELSE 'False' END AS [HasLimitingCollection],
    col.LimitToCollectionID AS [LimitingCollectionID],
    col.LimitToCollectionName AS [LimitingCollectionName],
    NULL AS [PackageID],
    NULL AS [PackageName],
    NULL AS [PackageType],
    NULL AS [ReferencePackageID],
    NULL AS [ReferencedPackageName],
    NULL AS [CollectionName],
    NULL AS [CollectionType],
    NULL AS [CollectionObjectPath]
FROM dbo.v_Collections col WITH (NOLOCK)
LEFT JOIN dbo.vFolderMembers mem WITH (NOLOCK)
    ON col.SiteID = mem.InstanceKey
LEFT JOIN FolderHierarchy hry
    ON mem.ContainerNodeID = hry.ContainerNodeID
LEFT JOIN dbo.v_Advertisement adv WITH (NOLOCK)
    ON col.SiteID = adv.CollectionID
WHERE col.CollectionType = 2
    AND mem.ObjectType = '5000'
    AND hry.[Path] LIKE '%Retired%'
    AND adv.AdvertisementID IS NULL

UNION ALL

-- Get Task Sequences in Retired folders WITH referenced packages and collections
SELECT DISTINCT
    'TaskSequence' AS [ObjectType],
    CAST(ISNULL(coll.CollectionID, '') AS nvarchar(50)) AS [CollectionID],
    CAST(ncol.CollectionID AS nvarchar(50)) AS [CollectionNumericID],
    tsp.PackageID AS [CollectionSiteID],
    tsp.Name AS [Name],
    ISNULL(ts_hry.[Path], '/') AS [FolderPath],
    0 AS [MemberCount],
    tsp.SourceDate AS [LastMemberChangeTime],
    DATEDIFF(day, tsp.SourceDate, GETDATE()) AS [DaysSinceLastChange],
    'False' AS [HasLimitingCollection],
    NULL AS [LimitingCollectionID],
    NULL AS [LimitingCollectionName],
    tsp.PackageID AS [PackageID],
    tsp.Name AS [PackageName],
    'Task Sequence Package' AS [PackageType],
    tsri.ReferencePackageID AS [ReferencePackageID],
    pkg.Name AS [ReferencedPackageName],
    coll.Name AS [CollectionName],
    CASE WHEN coll.Name IS NULL THEN NULL
        WHEN coll.CollectionType = 1 THEN 'UserCollection'
        ELSE 'DeviceCollection'
    END AS [CollectionType],
    ISNULL(coll_hry.[Path], '/') AS [CollectionObjectPath]
FROM v_TaskSequencePackage tsp WITH (NOLOCK)
LEFT JOIN v_TaskSequenceReferencesInfo tsri WITH (NOLOCK)
    ON tsp.PackageID = tsri.PackageID
LEFT JOIN v_Package pkg WITH (NOLOCK)
    ON tsri.ReferencePackageID = pkg.PackageID
LEFT JOIN v_Advertisement adv WITH (NOLOCK)
    ON tsp.PackageID = adv.PackageID
LEFT JOIN v_Collection coll WITH (NOLOCK)
    ON adv.CollectionID = coll.CollectionID
LEFT JOIN dbo.v_Collections ncol WITH (NOLOCK)
    ON ncol.SiteID = coll.CollectionID
LEFT JOIN vFolderMembers ts_mem WITH (NOLOCK)
    ON tsp.PackageID = ts_mem.InstanceKey
LEFT JOIN FolderHierarchy ts_hry
    ON ts_mem.ContainerNodeID = ts_hry.ContainerNodeID
LEFT JOIN vFolderMembers coll_mem WITH (NOLOCK)
    ON coll.CollectionID = coll_mem.InstanceKey
LEFT JOIN FolderHierarchy coll_hry
    ON coll_mem.ContainerNodeID = coll_hry.ContainerNodeID
WHERE ts_hry.[Path] LIKE '%Retired%'

ORDER BY [ObjectType], [Name]
"@

        if ($Progress) {
            $Progress.UpdateProgress("ProcessSites",
                "Site $SiteIndex/$TotalSites - Scanning retired folders",
                (($SiteIndex - 1) / $TotalSites * 100) + 5)
        }

        $dataset = New-Object System.Data.DataSet
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($retiredQuery, $SCCMConnection.Connection)
        $adapter.SelectCommand.CommandTimeout = 300  # 5 minute timeout

        $ResourceManager.Register($adapter, "RetiredFolderAdapter")

        [void]$adapter.Fill($dataset)

        $dataTable = $dataset.Tables[0]

        $Logger.Info("Retired folder scan completed. Objects found: $($dataTable.Rows.Count)", "RetiredFolderScan")

        # Log breakdown by type
        $userCollections = ($dataTable.Select("ObjectType = 'UserCollection'")).Count
        $deviceCollections = ($dataTable.Select("ObjectType = 'DeviceCollection'")).Count
        $taskSequences = ($dataTable.Select("ObjectType = 'TaskSequence'")).Count

        $Logger.Info("Breakdown - User Collections: $userCollections, Device Collections: $deviceCollections, Task Sequences: $taskSequences", "RetiredFolderScan")

        return @{
            Success = $true
            Data = $dataTable
            UserCollections = $userCollections
            DeviceCollections = $deviceCollections
            TaskSequences = $taskSequences
        }
    }
    catch {
        $ErrorTracker.RecordError("RetiredFolderScan",
            "Failed to scan retired folders",
            $false,
            $_.Exception)
        $Logger.Error("Retired folder scan failed: $_", "RetiredFolderScan")

        return @{
            Success = $false
            Data = $null
            UserCollections = 0
            DeviceCollections = 0
            TaskSequences = 0
        }
    }
}
#endregion

#region Retired Object Processor with EPP Validation
function Invoke-RetiredObjectProcessing {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Data.DataTable]$RetiredData,

        [Parameter(Mandatory = $true)]
        [string]$SiteId,

        [Parameter(Mandatory = $true)]
        [System.Data.SqlClient.SqlCommand]$MonitoringCommand,

        [Parameter(Mandatory = $true)]
        [Logger]$Logger,

        [Parameter(Mandatory = $true)]
        [ErrorTracker]$ErrorTracker,

        [Parameter(Mandatory = $true)]
        [datetime]$Timestamp,

        # NEW: EPP Data for validation
        [Parameter(Mandatory = $false)]
        [System.Data.DataTable]$EPPData = $null,

        # NEW: Task Sequence data for cross-reference validation
        [Parameter(Mandatory = $false)]
        [System.Data.DataTable]$TaskSequenceData = $null,

        [Parameter(Mandatory = $false)]
        [ProgressTracker]$Progress,

        [Parameter(Mandatory = $false)]
        [int]$SiteIndex = 1,

        [Parameter(Mandatory = $false)]
        [int]$TotalSites = 1
    )

    try {
        $Logger.Info("Starting retired object processing with EPP validation - Total objects: $($RetiredData.Rows.Count)", "RetiredObjectProcessor")

        $processedCount = 0
        $errorCount = 0
        $totalObjects = $RetiredData.Rows.Count
        $currentObject = 0

        # Separate retired data by type for better processing
        $retiredCollections = $RetiredData.Select("ObjectType = 'UserCollection' OR ObjectType = 'DeviceCollection'")
        $retiredTaskSequences = $RetiredData.Select("ObjectType = 'TaskSequence'")

        $Logger.Info("Processing breakdown - Collections (raw): $($retiredCollections.Count), Task Sequences (raw): $($retiredTaskSequences.Count)", "RetiredObjectProcessor")

        # Process Collections first
        foreach ($retiredItem in $retiredCollections) {
            $currentObject++

            # Update progress every 10 objects
            if ($Progress -and ($currentObject % 10 -eq 0 -or $currentObject -eq $totalObjects)) {
                $objectProgress = ($currentObject / $totalObjects) * 10  # Retired objects are 10% of site processing
                $siteBaseProgress = (($SiteIndex - 1) / $TotalSites * 100) + 5
                $Progress.UpdateProgress("ProcessSites",
                    "Site $SiteIndex/$TotalSites - Processing retired collections $currentObject of $totalObjects",
                    $siteBaseProgress + $objectProgress)
            }

            try {
                # NEW: Pass EPP data and Task Sequence data for validation
                $success = Add-RetiredCollectionToDatabase `
                    -RetiredItem $retiredItem `
                    -SiteId $SiteId `
                    -MonitoringCommand $MonitoringCommand `
                    -Logger $Logger `
                    -ErrorTracker $ErrorTracker `
                    -Timestamp $Timestamp `
                    -EPPData $EPPData `
                    -TaskSequenceData $TaskSequenceData

                if ($success) {
                    $processedCount++
                    $Logger.Debug("Processed retired collection: $($retiredItem.Name)", "RetiredObjectProcessor")
                } else {
                    $errorCount++
                }
            }
            catch {
                $errorCount++
                $ErrorTracker.RecordError("RetiredObjectProcessor",
                    "Error processing retired collection: $($retiredItem.Name)",
                    $false,
                    $_.Exception)
                $Logger.Error("Error processing retired collection $($retiredItem.Name): $_", "RetiredObjectProcessor")
            }
        }

        # Deduplicate task sequences by PackageID, keeping the row with the WORST EPP status so that
        # one active/unverified referenced package cannot be hidden behind a retired one.
        $statusRank = {
            param($row)
            $name = $row["ReferencedPackageName"]
            if ($null -eq $EPPData -or $name -is [DBNull] -or [string]::IsNullOrWhiteSpace($name)) { return 0 }
            $s = (Get-EPPPackageStatus -EPPData $EPPData -PackageName $name.ToString() -AppComments "" -Logger $Logger).Status
            if ($s -in @("Retired", "Cancel Package")) { return 1 }
            if ($s -in @("Not Found", "Null", "Error")) { return 2 }
            return 3
        }

        $uniqueTaskSequences = @{}
        $uniqueRank = @{}
        $tsCollectionIds = @{}
        foreach ($retiredItem in $retiredTaskSequences) {
            $packageId = $retiredItem.PackageID

            # Deduplication keeps one row per task sequence, so remember every collection it targets
            $numericId = $retiredItem["CollectionNumericID"]
            if ($numericId -isnot [DBNull] -and -not [string]::IsNullOrWhiteSpace($numericId)) {
                if (-not $tsCollectionIds.ContainsKey($packageId)) { $tsCollectionIds[$packageId] = @{} }
                $tsCollectionIds[$packageId][$numericId.ToString().Trim()] = $true
            }

            $rank = & $statusRank $retiredItem
            if (-not $uniqueTaskSequences.ContainsKey($packageId) -or $rank -gt $uniqueRank[$packageId]) {
                $uniqueTaskSequences[$packageId] = $retiredItem
                $uniqueRank[$packageId] = $rank
            }
        }

        $Logger.Info("Deduplicated task sequences - Raw count: $($retiredTaskSequences.Count), Unique count: $($uniqueTaskSequences.Count)", "RetiredObjectProcessor")

        # Process Task Sequences (using deduplicated list)
        foreach ($retiredItem in $uniqueTaskSequences.Values) {
            $currentObject++

            # Update progress every 10 objects
            if ($Progress -and ($currentObject % 10 -eq 0 -or $currentObject -eq $totalObjects)) {
                $objectProgress = ($currentObject / $totalObjects) * 10
                $siteBaseProgress = (($SiteIndex - 1) / $TotalSites * 100) + 5
                $Progress.UpdateProgress("ProcessSites",
                    "Site $SiteIndex/$TotalSites - Processing retired task sequences $currentObject of $totalObjects",
                    $siteBaseProgress + $objectProgress)
            }

            try {
                # NEW: Pass EPP data for validation
                $linkedCollectionIds = if ($tsCollectionIds.ContainsKey($retiredItem.PackageID)) {
                    @($tsCollectionIds[$retiredItem.PackageID].Keys)
                } else { @() }

                $success = Add-RetiredTaskSequenceToDatabase `
                    -RetiredItem $retiredItem `
                    -SiteId $SiteId `
                    -MonitoringCommand $MonitoringCommand `
                    -Logger $Logger `
                    -ErrorTracker $ErrorTracker `
                    -Timestamp $Timestamp `
                    -EPPData $EPPData `
                    -CollectionIds $linkedCollectionIds

                if ($success) {
                    $processedCount++
                    $Logger.Debug("Processed retired task sequence: $($retiredItem.Name)", "RetiredObjectProcessor")
                } else {
                    $errorCount++
                }
            }
            catch {
                $errorCount++
                $ErrorTracker.RecordError("RetiredObjectProcessor",
                    "Error processing retired task sequence: $($retiredItem.Name)",
                    $false,
                    $_.Exception)
                $Logger.Error("Error processing retired task sequence $($retiredItem.Name): $_", "RetiredObjectProcessor")
            }
        }

        $Logger.Info("Retired object processing completed - Processed: $processedCount, Errors: $errorCount", "RetiredObjectProcessor")

        # Log summary if EPP validation was performed
        if ($null -ne $EPPData) {
            $Logger.Info("EPP validation was performed on retired objects", "RetiredObjectProcessor")
        } else {
            $Logger.Warning("EPP validation was skipped - no EPP data available", "RetiredObjectProcessor")
        }

        # FIX: Return deduplicated counts for accurate reporting
        return @{
            Success = $true
            ProcessedCount = $processedCount
            ErrorCount = $errorCount
            # NEW: Return actual deduplicated counts
            UniqueCollectionsCount = $retiredCollections.Count
            UniqueTaskSequencesCount = $uniqueTaskSequences.Count
        }
    }
    catch {
        $ErrorTracker.RecordError("RetiredObjectProcessor",
            "Fatal error during retired object processing",
            $true,
            $_.Exception)
        $Logger.Fatal("Retired object processing failed: $_", "RetiredObjectProcessor")

        return @{
            Success = $false
            ProcessedCount = 0
            ErrorCount = $errorCount
            UniqueCollectionsCount = 0
            UniqueTaskSequencesCount = 0
        }
    }
}
#endregion

#region Task Sequence EPP Cleanup Orchestrator
function Invoke-TaskSequenceEPPCleanup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Data.DataTable]$TaskSequenceData,

        [Parameter(Mandatory = $true)]
        [System.Data.DataTable]$EPPData,

        [Parameter(Mandatory = $true)]
        [string]$SiteId,

        [Parameter(Mandatory = $true)]
        [System.Data.SqlClient.SqlCommand]$MonitoringCommand,

        [Parameter(Mandatory = $true)]
        [Logger]$Logger,

        [Parameter(Mandatory = $true)]
        [ErrorTracker]$ErrorTracker,

        [Parameter(Mandatory = $true)]
        [datetime]$Timestamp,

        # NEW PARAMETER: Retired folder data to exclude from EPP filtering
        [Parameter(Mandatory = $false)]
        [System.Data.DataTable]$RetiredFolderData = $null
    )

    try {
        $Logger.Info("Starting Task Sequence EPP cleanup processing for SiteId: $SiteId", "TSEPPCleanup")

        # Build EPP lookup tables
        $eppLookup = Get-EPPLookup -EPPData $EPPData -Logger $Logger

        # Get Task Sequence EPP matches
        $tsMatches = Get-TaskSequenceEPPMatches `
            -TaskSequenceData $TaskSequenceData `
            -Lookup $eppLookup `
            -Logger $Logger

        # NEW: Build exclusion list of Package IDs already in retired folders
        $retiredFolderPackageIds = @()
        if ($null -ne $RetiredFolderData -and $RetiredFolderData.Rows.Count -gt 0) {
            $retiredFolderPackageIds = $RetiredFolderData.Rows |
                Where-Object { $_.ObjectType -eq 'TaskSequence' -and -not [string]::IsNullOrWhiteSpace($_.PackageID) } |
                ForEach-Object { $_.PackageID.ToString().Trim() }

            $Logger.Info("Excluding $($retiredFolderPackageIds.Count) task sequences already in retired folders from EPP filtering", "TSEPPCleanup")
        }

        $retiredTS = Select-RetiredTaskSequenceMatches -Matches $tsMatches -ExcludePackageIds $retiredFolderPackageIds

        $Logger.Info("Found $($retiredTS.Count) task sequences with retired/cancelled EPP packages (excluding those in retired folders)", "TSEPPCleanup")

        $processedCount = 0
        $errorCount = 0

        # Process each retired task sequence
        foreach ($tsItem in $retiredTS) {
            $success = Add-TaskSequenceEPPCleanup `
                -TaskSequenceItem $tsItem `
                -SiteId $SiteId `
                -MonitoringCommand $MonitoringCommand `
                -Logger $Logger `
                -ErrorTracker $ErrorTracker `
                -Timestamp $Timestamp

            if ($success) {
                $processedCount++
            } else {
                $errorCount++
            }
        }

        $Logger.Info("Task Sequence EPP cleanup completed - Processed: $processedCount, Errors: $errorCount", "TSEPPCleanup")

        return @{
            Success = $true
            ProcessedCount = $processedCount
            ErrorCount = $errorCount
        }
    }
    catch {
        $ErrorTracker.RecordError("TSEPPCleanup",
            "Failed to process Task Sequence EPP cleanup",
            $false,
            $_.Exception)
        $Logger.Error("Task Sequence EPP cleanup failed: $_", "TSEPPCleanup")

        return @{
            Success = $false
            ProcessedCount = 0
            ErrorCount = 0
        }
    }
}
#endregion

#region Site Registration
function Register-SCCMSite {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$SiteItem,

        [Parameter(Mandatory = $true)]
        [string]$Environment,

        [Parameter(Mandatory = $true)]
        [System.Data.SqlClient.SqlConnection]$MonitoringConnection,

        [Parameter(Mandatory = $true)]
        [Logger]$Logger,

        [Parameter(Mandatory = $true)]
        [ErrorTracker]$ErrorTracker,

        [Parameter(Mandatory = $true)]
        [datetime]$Timestamp  # CHANGED from [string] to [datetime]
    )

    try {
        $siteName = $SiteItem.site
        $Logger.Info("Registering SCCM site: $siteName", "SiteRegistration")

        # FIXED: Use parameterized query for INSERT
        $sqlInsert = @"
INSERT INTO [dbo].[tbl_SCCMSites] ([site],[environment],[timestamp])
SELECT @Site, @Environment, @Timestamp
WHERE NOT EXISTS (
    SELECT [site],[environment] FROM [dbo].[tbl_SCCMSites]
    WHERE [site] = @Site AND [environment] = @Environment
)
"@

        $cmd = $MonitoringConnection.CreateCommand()
        $cmd.Parameters.Clear()
        [void]$cmd.Parameters.AddWithValue("@Site", $SiteItem.site)
        [void]$cmd.Parameters.AddWithValue("@Environment", $Environment)

        # Add timestamp as DateTime object
        $timestampParam = $cmd.Parameters.Add("@Timestamp", [System.Data.SqlDbType]::DateTime)
        $timestampParam.Value = $Timestamp

        $cmd.CommandText = $sqlInsert
        [void]$cmd.ExecuteNonQuery()
        $cmd.Dispose()

        # Use parameterized query for SELECT
        $sqlSelect = @"
SELECT TOP 1 s1.[id]
FROM [dbo].[tbl_SCCMSites] s1 WITH (NOLOCK)
WHERE s1.[site] = @Site
    AND s1.[environment] = @Environment
    AND s1.[timestamp] = (
        SELECT MAX(s2.[timestamp])
        FROM [dbo].[tbl_SCCMSites] s2 WITH (NOLOCK)
        WHERE s2.[site] = s1.[site]
            AND s2.[environment] = s1.[environment]
    )
"@

        $cmd = $MonitoringConnection.CreateCommand()
        $cmd.Parameters.Clear()
        [void]$cmd.Parameters.AddWithValue("@Site", $SiteItem.site)
        [void]$cmd.Parameters.AddWithValue("@Environment", $Environment)

        $cmd.CommandText = $sqlSelect

        $reader = $cmd.ExecuteReader()
        $siteId = $null

        while ($reader.Read()) {
            $siteId = $reader.GetValue(0)
        }

        $reader.Close()
        $cmd.Dispose()

        if ($null -eq $siteId) {
            throw "Failed to retrieve site ID for: $siteName"
        }

        $Logger.Info("Site registered successfully with ID: $siteId", "SiteRegistration")

        return @{
            Success = $true
            SiteId = $siteId
        }
    }
    catch {
        $ErrorTracker.RecordError("SiteRegistration",
            "Failed to register site: $($SiteItem.site)",
            $false,
            $_.Exception)
        $Logger.Error("Site registration failed: $_", "SiteRegistration")

        return @{
            Success = $false
            SiteId = $null
        }
    }
}
#endregion

#region Site Data Refresh
function Clear-SiteMonitoringData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteId,

        [Parameter(Mandatory = $true)]
        [System.Data.SqlClient.SqlCommand]$MonitoringCommand,

        [Parameter(Mandatory = $true)]
        [Logger]$Logger,

        [Parameter(Mandatory = $true)]
        [ErrorTracker]$ErrorTracker
    )

    # Removes this site's previous results in one transaction. Rows without a siteid are matched
    # through the site's collections and task sequence package IDs.
    $sqlClear = @"
SET XACT_ABORT ON;
BEGIN TRAN;
DELETE FROM [dbo].[tbl_SCCMapplications]
WHERE [collectionid] IN (SELECT [collectionid] FROM [dbo].[tbl_SCCMCollections] WHERE [siteid] = @SiteId)
   OR [collectionsiteid] IN (SELECT [Task Sequence Package ID] FROM [dbo].[tbl_SCCMTaskSequencesEPPCleanup] WHERE [siteid] = @SiteId);
DELETE FROM [dbo].[tbl_SCCMtasksequences]
WHERE [collectionid] IN (SELECT [collectionid] FROM [dbo].[tbl_SCCMCollections] WHERE [siteid] = @SiteId)
   OR [collectionsiteid] IN (SELECT [Task Sequence Package ID] FROM [dbo].[tbl_SCCMTaskSequencesEPPCleanup] WHERE [siteid] = @SiteId);
DELETE FROM [dbo].[tbl_SCCMCollectionStatus]
WHERE [collectionid] IN (SELECT [collectionid] FROM [dbo].[tbl_SCCMCollections] WHERE [siteid] = @SiteId);
DELETE FROM [dbo].[tbl_SCCMCollections] WHERE [siteid] = @SiteId;
DELETE FROM [dbo].[tbl_SCCMTaskSequencesEPPCleanup] WHERE [siteid] = @SiteId;
COMMIT;
"@

    try {
        $MonitoringCommand.Parameters.Clear()
        [void]$MonitoringCommand.Parameters.AddWithValue("@SiteId", $SiteId)
        $MonitoringCommand.CommandText = $sqlClear
        $removed = $MonitoringCommand.ExecuteNonQuery()
        $Logger.Info("Replaced previous results for SiteId $SiteId - rows removed: $removed", "SiteRefresh")
        return $true
    }
    catch {
        $ErrorTracker.RecordError("SiteRefresh",
            "Failed to clear previous results for SiteId $SiteId",
            $false,
            $_.Exception)
        $Logger.Error("Failed to clear previous results for SiteId ${SiteId}: $_", "SiteRefresh")
        return $false
    }
}
#endregion

#region Collection Processing Orchestrator
function Invoke-CollectionProcessing {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Data.DataTable]$CollectionData,

        [Parameter(Mandatory = $true)]
        [System.Data.DataTable]$EPPData,

        [Parameter(Mandatory = $true)]
        [string]$SiteId,

        [Parameter(Mandatory = $true)]
        [System.Data.SqlClient.SqlCommand]$MonitoringCommand,

        [Parameter(Mandatory = $true)]
        [Logger]$Logger,

        [Parameter(Mandatory = $true)]
        [ErrorTracker]$ErrorTracker,

        [Parameter(Mandatory = $true)]
        [datetime]$Timestamp,

        [Parameter(Mandatory = $false)]
        [ProgressTracker]$Progress,

        [Parameter(Mandatory = $false)]
        [int]$SiteIndex = 1,

        [Parameter(Mandatory = $false)]
        [int]$TotalSites = 1
    )

    try
    {
        $Logger.Info("Starting collection processing - Total collections: $($CollectionData.Rows.Count)", "CollectionProcessor")

        $processedCount = 0
        $skippedCount = 0
        $errorCount = 0
        $totalCollections = $CollectionData.Rows.Count
        $currentCollection = 0

        # Pre-pass: collection rows are one per application/deployment, so a collection's
        # EPP status must be judged across ALL of its rows, not the row being iterated.
        $rowEppStatus = New-Object 'System.Collections.Generic.List[object]'
        $collectionEppStatuses = @{}
        foreach ($row in $CollectionData.Rows) {
            $rowPackage = if ([string]::IsNullOrWhiteSpace($row.ApplicationName)) { "N/A" } else { $row.ApplicationName.ToString().Trim() }
            $rowComments = if ([string]::IsNullOrWhiteSpace($row.App_Comments)) { "" } else { $row.App_Comments.ToString().Trim() }
            try {
                $rowStatus = Get-EPPPackageStatus -EPPData $EPPData -PackageName $rowPackage -AppComments $rowComments -Logger $Logger
            }
            catch {
                $rowStatus = @{ Status = "Error"; Bypass = $true; MatchType = "Error" }
            }
            $rowEppStatus.Add($rowStatus)

            $key = $row.CollectionID.ToString()
            if (-not $collectionEppStatuses.ContainsKey($key)) {
                $collectionEppStatuses[$key] = New-Object 'System.Collections.Generic.List[object]'
            }
            $collectionEppStatuses[$key].Add($rowStatus)
        }

        $adGroupCache = @{}
        foreach ($collectionItem in $CollectionData.Rows) {
            $currentCollection++

            # Update progress every 10 collections or on last collection
            if ($Progress -and ($currentCollection % 10 -eq 0 -or $currentCollection -eq $totalCollections)) {
                $collectionProgress = ($currentCollection / $totalCollections) * 40  # Collections are 40% of site processing (70-100%)
                $siteBaseProgress = (($SiteIndex - 1) / $TotalSites * 100) + 60
                $Progress.UpdateProgress("ProcessSites",
                    "Site $SiteIndex/$TotalSites - Processing collections $currentCollection of $totalCollections",
                    $siteBaseProgress + $collectionProgress)
            }

            try
            {
                $rowStatus = $rowEppStatus[$currentCollection - 1]

                # The row keeps its own status for the monitoring record
                if ($collectionItem.PSObject.Properties.Name -contains 'eppstatus') {
                    $collectionItem.eppstatus = $rowStatus.Status
                } else {
                    $collectionItem | Add-Member -NotePropertyName "eppstatus" -NotePropertyValue $rowStatus.Status -Force
                }

                # The decision uses the whole collection: any active/unverifiable package blocks it,
                # and Retired applies only when every named package is Retired/Cancel Package.
                $namedStatuses = @($collectionEppStatuses[$collectionItem.CollectionID.ToString()] |
                    Where-Object { $_.MatchType -ne "No Package Name" })
                $blocking = @($namedStatuses | Where-Object { $_.Status -notin @("Retired", "Cancel Package", "Not Found", "Null") })

                if ($blocking.Count -gt 0) {
                    $eppStatus = @{ Status = $blocking[0].Status; Bypass = $true; MatchType = "Collection - active or unverified package" }
                }
                elseif ($namedStatuses.Count -gt 0 -and
                        @($namedStatuses | Where-Object { $_.Status -in @("Retired", "Cancel Package") }).Count -eq $namedStatuses.Count) {
                    $eppStatus = @{ Status = $namedStatuses[0].Status; Bypass = $false; MatchType = "Collection - all packages retired" }
                }
                else {
                    $eppStatus = @{ Status = "Not Found"; Bypass = $true; MatchType = "Collection - no confirmed EPP match" }
                }
            }
            catch {
                $Logger.Warning("Error getting EPP status for $($collectionItem.CollectionName): $_", "CollectionProcessor")
                $eppStatus = @{
                    Status = "Error"
                    Bypass = $true
                    MatchType = "Error"
                }
            }

            try {
            # Get AD group information (cached per collection)
            $adKey = $collectionItem.CollectionID.ToString()
            if (-not $adGroupCache.ContainsKey($adKey)) {
                $adGroupCache[$adKey] = Get-ADGroupInformation `
                    -CollectionName $collectionItem.CollectionName `
                    -CollectionType $collectionItem.CollectionType `
                    -ReferenceDate (Get-Date $Timestamp) `
                    -Logger $Logger `
                    -ErrorTracker $ErrorTracker
            }
            $adGroupInfo = $adGroupCache[$adKey]
            # Test processing rules
            $processingDecision = Test-CollectionProcessingRules `
                -CollectionItem $collectionItem `
                -EPPStatus $eppStatus `
                -ADGroupInfo $adGroupInfo `
                -ReferenceDate (Get-Date $Timestamp) `
                -Logger $Logger

            # Process if rules are met
            if ($processingDecision.ShouldProcess) {
                $Logger.Info("Processing collection: $($collectionItem.CollectionName) - Reason: $($processingDecision.Reason)", "CollectionProcessor")

                # Determine which date to use
                $whenChanged = if ($processingDecision.UseADGroupDate) {
                    $adGroupInfo.WhenChangedString
                } else {
                    $null
                }

                $daysSince = if ($processingDecision.UseADGroupDate) {
                    $adGroupInfo.DaysSinceChange
                } else {
                    0
                }

                # Insert collection data
                $success = Add-CollectionToDatabase `
                    -CollectionItem $collectionItem `
                    -SiteId $SiteId `
                    -MonitoringCommand $MonitoringCommand `
                    -Logger $Logger `
                    -ErrorTracker $ErrorTracker `
                    -Timestamp $Timestamp `
                    -WhenChangedString $whenChanged `
                    -DaysSinceChange $daysSince

                if ($success) {
                    # Insert application data
                    Add-ApplicationToDatabase `
                        -CollectionItem $collectionItem `
                        -MonitoringCommand $MonitoringCommand `
                        -Logger $Logger `
                        -ErrorTracker $ErrorTracker `
                        -Timestamp $Timestamp | Out-Null

                    # Insert task sequence data
                    Add-TaskSequenceToDatabase `
                        -CollectionItem $collectionItem `
                        -MonitoringCommand $MonitoringCommand `
                        -Logger $Logger `
                        -ErrorTracker $ErrorTracker `
                        -Timestamp $Timestamp | Out-Null

                    # Insert collection status
                    Add-CollectionStatusToDatabase `
                        -CollectionId $collectionItem.CollectionID `
                        -MonitoringCommand $MonitoringCommand `
                        -Logger $Logger `
                        -ErrorTracker $ErrorTracker `
                        -Timestamp $Timestamp | Out-Null

                    $processedCount++
                } else {
                    $errorCount++
                }
            } else {
                $skippedCount++
                $Logger.Debug("Skipped collection: $($collectionItem.CollectionName) - $($processingDecision.Reason)", "CollectionProcessor")
            }
        }
        catch {
            $errorCount++
            $ErrorTracker.RecordError("CollectionProcessor",
                "Error processing collection: $($collectionItem.CollectionName)",
                $false,
                $_.Exception)
            $Logger.Error("Error processing collection $($collectionItem.CollectionName): $_", "CollectionProcessor")
        }
    }

    $Logger.Info("Collection processing completed - Processed: $processedCount, Skipped: $skippedCount, Errors: $errorCount", "CollectionProcessor")

    return @{
        Success = $true
        ProcessedCount = $processedCount
        SkippedCount = $skippedCount
        ErrorCount = $errorCount
    }
}
catch {
    $ErrorTracker.RecordError("CollectionProcessor",
        "Fatal error during collection processing",
        $true,
        $_.Exception)
    $Logger.Fatal("Collection processing failed: $_", "CollectionProcessor")

    return @{
        Success = $false
        ProcessedCount = 0
        SkippedCount = 0
        ErrorCount = 0
    }
}
}
#endregion

#region Site Processing Orchestrator
function Invoke-SiteProcessing {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$SiteItem,

        [Parameter(Mandatory = $true)]
        [System.Data.DataTable]$EPPData,

        [Parameter(Mandatory = $true)]
        [string]$Environment,

        [Parameter(Mandatory = $true)]
        [DatabaseConnection]$MonitoringConnection,

        [Parameter(Mandatory = $true)]
        [Logger]$Logger,
[Parameter(Mandatory = $true)]
[ErrorTracker]$ErrorTracker,

[Parameter(Mandatory = $true)]
[ResourceManager]$ResourceManager,

[Parameter(Mandatory = $true)]
[datetime]$Timestamp,

[Parameter(Mandatory = $false)]
[ProgressTracker]$Progress,

[Parameter(Mandatory = $false)]
[int]$SiteIndex = 1,

[Parameter(Mandatory = $false)]
[int]$TotalSites = 1
)

$sccmConnection = $null

try {
    $Logger.Info("Starting processing for site: $($SiteItem.site)", "SiteProcessor")

    # Connect to SCCM site (10% of site processing)
    if ($Progress) {
        $Progress.UpdateProgress("ProcessSites", "Site $SiteIndex/$TotalSites - Connecting to SCCM: $($SiteItem.site)",
            (($SiteIndex - 1) / $TotalSites * 100) + 2)
    }

    $sccmResult = Connect-SCCMSite `
        -SiteItem $SiteItem `
        -Logger $Logger `
        -ErrorTracker $ErrorTracker `
        -ResourceManager $ResourceManager

    if (-not $sccmResult.Success) {
        throw "Failed to connect to SCCM site: $($SiteItem.site)"
    }

    $sccmConnection = $sccmResult.Connection

    # Execute Task Sequence query (20% of site processing)
    if ($Progress) {
        $Progress.UpdateProgress("ProcessSites", "Site $SiteIndex/$TotalSites - Querying Task Sequences: $($SiteItem.site)",
            (($SiteIndex - 1) / $TotalSites * 100) + 10)
    }

    $Logger.Info("Executing Task Sequence query for site: $($SiteItem.site)", "SiteProcessor")

    $tsDataResult = Get-TaskSequenceData `
        -SCCMConnection $sccmConnection `
        -Logger $Logger `
        -ErrorTracker $ErrorTracker `
        -ResourceManager $ResourceManager

    # Execute Collection Detection query
    if ($Progress) {
        $Progress.UpdateProgress("ProcessSites",
            "Site $SiteIndex/$TotalSites - Detecting Collections: $($SiteItem.site)",
            (($SiteIndex - 1) / $TotalSites * 100) + 30)
    }

    $Logger.Info("Executing Collection Detection query for site: $($SiteItem.site)", "SiteProcessor")

    $collectionDataResult = Invoke-CollectionDetectionQuery `
        -SCCMConnection $sccmConnection `
        -AdaptivaSqlDbName $SiteItem.adaptivasqldbname `
        -Logger $Logger `
        -ErrorTracker $ErrorTracker `
        -ResourceManager $ResourceManager `
        -Progress $Progress `
        -SiteIndex $SiteIndex `
        -TotalSites $TotalSites

    if (-not $collectionDataResult.Success) {
        throw "Collection Detection query failed for site: $($SiteItem.site)"
    }

    # Execute Retired Folder Scan (15% of site processing)
    if ($Progress) {
        $Progress.UpdateProgress("ProcessSites",
            "Site $SiteIndex/$TotalSites - Scanning retired folders: $($SiteItem.site)",
            (($SiteIndex - 1) / $TotalSites * 100) + 55)
    }

    $Logger.Info("Scanning retired folders for site: $($SiteItem.site)", "SiteProcessor")

    $retiredDataResult = Invoke-RetiredFolderScan `
        -SCCMConnection $sccmConnection `
        -Logger $Logger `
        -ErrorTracker $ErrorTracker `
        -ResourceManager $ResourceManager `
        -Progress $Progress `
        -SiteIndex $SiteIndex `
        -TotalSites $TotalSites

    # Register site in monitoring database (60% of site processing)
    if ($Progress) {
        $Progress.UpdateProgress("ProcessSites", "Site $SiteIndex/$TotalSites - Registering site: $($SiteItem.site)",
            (($SiteIndex - 1) / $TotalSites * 100) + 60)
    }

    $siteIdResult = Register-SCCMSite `
        -SiteItem $SiteItem `
        -Environment $Environment `
        -MonitoringConnection $MonitoringConnection.Connection `
        -Logger $Logger `
        -ErrorTracker $ErrorTracker `
        -Timestamp $Timestamp

    if (-not $siteIdResult.Success) {
        throw "Failed to register site: $($SiteItem.site)"
    }

    # All SCCM queries have succeeded, so it is now safe to replace this site's previous results
    if ($script:KeepHistory) {
        $Logger.Info("KeepHistory specified - previous results for site $($SiteItem.site) are kept", "SiteProcessor")
    }
    elseif (-not $retiredDataResult.Success -or -not $tsDataResult.Success) {
        $Logger.Warning("A source query failed - previous results for site $($SiteItem.site) are kept", "SiteProcessor")
    }
    else {
        $refreshed = Clear-SiteMonitoringData `
            -SiteId $siteIdResult.SiteId `
            -MonitoringCommand $MonitoringConnection.Command `
            -Logger $Logger `
            -ErrorTracker $ErrorTracker
        if (-not $refreshed) {
            throw "Failed to replace previous results for site: $($SiteItem.site)"
        }
    }

    # Initialize counters for detailed tracking
    $retiredCollectionsCount = 0
    $retiredTaskSequencesCount = 0

    # Process Retired Objects (step 1)
    $retiredProcessingResult = @{
        ProcessedCount = 0
        ErrorCount = 0
        UniqueCollectionsCount = 0
        UniqueTaskSequencesCount = 0
    }

    if ($retiredDataResult.Success -and $retiredDataResult.Data.Rows.Count -gt 0) {
        if ($Progress) {
            $Progress.UpdateProgress("ProcessSites",
                "Site $SiteIndex/$TotalSites - Processing retired objects: $($SiteItem.site)",
                (($SiteIndex - 1) / $TotalSites * 100) + 68)
        }

        $Logger.Info("Processing $($retiredDataResult.Data.Rows.Count) retired objects (raw count) for site: $($SiteItem.site)", "SiteProcessor")

        $retiredProcessingResult = Invoke-RetiredObjectProcessing `
            -RetiredData $retiredDataResult.Data `
            -SiteId $siteIdResult.SiteId `
            -MonitoringCommand $MonitoringConnection.Command `
            -Logger $Logger `
            -ErrorTracker $ErrorTracker `
            -Timestamp $Timestamp `
            -EPPData $EPPData `
            -TaskSequenceData $tsDataResult.Data `
            -Progress $Progress `
            -SiteIndex $SiteIndex `
            -TotalSites $TotalSites

        if ($retiredProcessingResult.Success) {
            # FIX: Use deduplicated counts from processing result
            $retiredCollectionsCount = $retiredProcessingResult.UniqueCollectionsCount
            $retiredTaskSequencesCount = $retiredProcessingResult.UniqueTaskSequencesCount

            $Logger.Info("Retired object processing completed - Collections: $retiredCollectionsCount, Task Sequences: $retiredTaskSequencesCount (deduplicated)", "SiteProcessor")
        }
    }
    else {
        $Logger.Info("No retired objects found for site: $($SiteItem.site)", "SiteProcessor")
    }

    # Process Staging Orphaned Objects (step 2)
    if ($Progress) {
        $Progress.UpdateProgress("ProcessSites",
            "Site $SiteIndex/$TotalSites - Scanning staging for orphaned objects: $($SiteItem.site)",
            (($SiteIndex - 1) / $TotalSites * 100) + 69)
    }

    $Logger.Info("Scanning staging folders for orphaned objects for site: $($SiteItem.site)", "SiteProcessor")

    $orphanedObjectsResult = Invoke-StagingOrphanedObjectsProcessing `
        -SCCMConnection $sccmConnection `
        -SiteId $siteIdResult.SiteId `
        -MonitoringCommand $MonitoringConnection.Command `
        -Logger $Logger `
        -ErrorTracker $ErrorTracker `
        -ResourceManager $ResourceManager `
        -Timestamp $Timestamp `
        -RetiredFolderData $retiredDataResult.Data `
        -Progress $Progress `
        -SiteIndex $SiteIndex `
        -TotalSites $TotalSites

    if ($orphanedObjectsResult.Success) {
        $Logger.Info("Staging orphaned objects scan completed - Device Collections: $($orphanedObjectsResult.DeviceCollections), " +
            "User Collections: $($orphanedObjectsResult.UserCollections), " +
            "Task Sequences: $($orphanedObjectsResult.TaskSequences)", "SiteProcessor")
    }
    else {
        $Logger.Warning("Staging orphaned objects scan encountered errors", "SiteProcessor")
    }

    # Process Task Sequences with EPP cleanup (step 3)
    $tsEppCleanupResult = @{ ProcessedCount = 0 }
    if ($tsDataResult.Success) {
        if ($Progress) {
            $Progress.UpdateProgress("ProcessSites", "Site $SiteIndex/$TotalSites - Processing TS EPP cleanup: $($SiteItem.site)",
                (($SiteIndex - 1) / $TotalSites * 100) + 65)
        }

        $Logger.Info("Processing Task Sequences for EPP cleanup with SiteId: $($siteIdResult.SiteId)", "SiteProcessor")

        $tsEppCleanupResult = Invoke-TaskSequenceEPPCleanup `
            -TaskSequenceData $tsDataResult.Data `
            -EPPData $EPPData `
            -SiteId $siteIdResult.SiteId `
            -MonitoringCommand $MonitoringConnection.Command `
            -Logger $Logger `
            -ErrorTracker $ErrorTracker `
            -Timestamp $Timestamp `
            -RetiredFolderData $retiredDataResult.Data
    }

    # Process collections (step 4)
    if ($Progress) {
        $Progress.UpdateProgress("ProcessSites", "Site $SiteIndex/$TotalSites - Processing Collections: $($SiteItem.site)",
            (($SiteIndex - 1) / $TotalSites * 100) + 70)
    }

    $processingResult = Invoke-CollectionProcessing `
        -CollectionData $collectionDataResult.Data `
        -EPPData $EPPData `
        -SiteId $siteIdResult.SiteId `
        -MonitoringCommand $MonitoringConnection.Command `
        -Logger $Logger `
        -ErrorTracker $ErrorTracker `
        -Timestamp $Timestamp `
        -Progress $Progress `
        -SiteIndex $SiteIndex `
        -TotalSites $TotalSites

    # Calculate orphaned collections count (device + user)
    $orphanedCollectionsCount = $orphanedObjectsResult.DeviceCollections + $orphanedObjectsResult.UserCollections

    $Logger.Info("Site processing completed for: $($SiteItem.site)", "SiteProcessor")
    $Logger.Info("Summary - Regular Collections: $($processingResult.ProcessedCount), " +
        "Retired Collections: $retiredCollectionsCount, " +
        "Retired Task Sequences: $retiredTaskSequencesCount, " +
        "Orphaned Collections: $orphanedCollectionsCount, " +
        "Orphaned Task Sequences: $($orphanedObjectsResult.TaskSequences), " +
        "EPP Cleanup Task Sequences: $($tsEppCleanupResult.ProcessedCount)", "SiteProcessor")

    # Return detailed breakdown with deduplicated counts
    return @{
        Success = $true
        SiteName = $SiteItem.site

        # Regular collections from collection detection query
        ProcessedCount = $processingResult.ProcessedCount

        # Task sequences from EPP cleanup (non-retired, non-orphaned)
        TaskSequencesProcessed = $tsEppCleanupResult.ProcessedCount

        # Retired objects (deduplicated counts)
        RetiredCollectionsProcessed = $retiredCollectionsCount
        RetiredTaskSequencesProcessed = $retiredTaskSequencesCount

        # Orphaned objects (separate counts)
        OrphanedCollectionsProcessed = $orphanedCollectionsCount
        OrphanedTaskSequencesProcessed = $orphanedObjectsResult.TaskSequences
    }
}
catch {
    $ErrorTracker.RecordError("SiteProcessor",
        "Fatal error processing site: $($SiteItem.site)",
        $true,
        $_.Exception)
    $Logger.Fatal("Site processing failed for $($SiteItem.site): $_", "SiteProcessor")

    return @{
        Success = $false
        SiteName = $SiteItem.site
        ProcessedCount = 0
        TaskSequencesProcessed = 0
        RetiredCollectionsProcessed = 0
        RetiredTaskSequencesProcessed = 0
        OrphanedCollectionsProcessed = 0
        OrphanedTaskSequencesProcessed = 0
    }
}
finally {
    # Cleanup SCCM connection
    if ($sccmConnection) {
        $sccmConnection.Disconnect($Logger)
    }
}
}
#endregion

#region Main Execution Orchestrator
function Invoke-SCCMCollectionScan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ScriptConfig]$Config,

        [Parameter(Mandatory = $true)]
        [Logger]$Logger,

        [Parameter(Mandatory = $true)]
        [ErrorTracker]$ErrorTracker,

        [Parameter(Mandatory = $true)]
        [ResourceManager]$ResourceManager
    )

    $timer = [System.Diagnostics.Stopwatch]::StartNew()

    # Initialize progress tracker
    $progress = [ProgressTracker]::new("SCCM Collection Scan", $Logger)
    $progress.DefineSteps(@{
        "Initialize" = 5
        "LoadConfig" = 5
        "ParseEnvironment" = 5
        "ConnectDatabases" = 15
        "LoadEPPData" = 10
        "ProcessSites" = 55
        "Finalize" = 5
    })

    try {
        $Logger.Info("=== SCCM Collection Scan Started ===", "MainOrchestrator")
        $Logger.Info("Environment: $($Config.Environment), Days Threshold: $($Config.Days)", "MainOrchestrator")

        # Step 1: Initialize directories
        $progress.StartStep("Initialize", "Creating directory structure")
        if (-not (Initialize-Directories -Config $Config -Logger $Logger -ErrorTracker $ErrorTracker)) {
            throw "Failed to initialize directories"
        }
        $progress.CompleteStep("Initialize")

        # Step 2: Load configuration
        $progress.StartStep("LoadConfig", "Loading XML configuration")
        $configResult = Get-ConfigurationData -Config $Config -Logger $Logger -ErrorTracker $ErrorTracker
        if (-not $configResult.Success) {
            throw "Failed to load configuration"
        }
        $progress.CompleteStep("LoadConfig")

        # Step 3: Parse environment configuration
        $progress.StartStep("ParseEnvironment", "Parsing environment settings")
        $envConfigResult = Get-EnvironmentConfiguration `
            -XmlData $configResult.XmlData `
            -Environment $Config.Environment `
            -Logger $Logger `
            -ErrorTracker $ErrorTracker

        if (-not $envConfigResult.Success) {
            throw "Failed to parse environment configuration"
        }
        $envConfig = $envConfigResult.Configuration
        $progress.CompleteStep("ParseEnvironment")

        # Step 4: Initialize database connections
        $progress.StartStep("ConnectDatabases", "Connecting to databases")
        $dbResult = Initialize-DatabaseConnections `
            -EnvConfig $envConfig `
            -Logger $Logger `
            -ErrorTracker $ErrorTracker `
            -ResourceManager $ResourceManager

        if (-not $dbResult.Success) {
            throw "Failed to initialize database connections"
        }
        $progress.CompleteStep("ConnectDatabases")

        # Step 5: Load EPP data (make it optional)
        $progress.StartStep("LoadEPPData", "Loading EPP package data")

        # Check if SPDW connection exists
        if ($null -eq $dbResult.SPDWConnection -or -not $dbResult.SPDWConnection.IsConnected) {
            $Logger.Warning("SPDW database not available - skipping EPP package data", "MainOrchestrator")

            # Create empty EPP data structure
            $eppDataResult = @{
                Success = $true
                Data = New-Object System.Data.DataTable
            }
            [void]$eppDataResult.Data.Columns.Add("EPP Status", [string])
            [void]$eppDataResult.Data.Columns.Add("Package Name", [string])
            [void]$eppDataResult.Data.Columns.Add("Package Short Name", [string])
        } else {
            $eppDataResult = Get-EPPPackageData `
                -SPDWConnection $dbResult.SPDWConnection `
                -Logger $Logger `
                -ErrorTracker $ErrorTracker `
                -ResourceManager $ResourceManager `
                -MaxRetries 3 `
                -RetryDelaySeconds 15

            if (-not $eppDataResult.Success) {
                $Logger.Warning("Failed to load EPP package data - continuing without it", "MainOrchestrator")

                # Create empty fallback
                $eppDataResult = @{
                    Success = $true
                    Data = New-Object System.Data.DataTable
                }
                [void]$eppDataResult.Data.Columns.Add("EPP Status", [string])
                [void]$eppDataResult.Data.Columns.Add("Package Name", [string])
                [void]$eppDataResult.Data.Columns.Add("Package Short Name", [string])
            }
        }

        $progress.UpdateProgress("LoadEPPData", "Loaded $($eppDataResult.Data.Rows.Count) EPP packages", 100)
        $progress.CompleteStep("LoadEPPData")

        # Step 6: Process each site
        $progress.StartStep("ProcessSites", "Processing SCCM sites")

        # Initialize detailed counters
        $totalRegularCollections = 0
        $totalRetiredCollections = 0
        $totalOrphanedCollections = 0
        $totalEppTaskSequences = 0
        $totalRetiredTaskSequences = 0
        $totalOrphanedTaskSequences = 0

        $sitesProcessed = 0
        $sitesFailed = 0
        $siteCount = $envConfig.Sites.Count
        $currentSiteIndex = 0

        foreach ($siteItem in $envConfig.Sites) {
            $currentSiteIndex++
            $siteProgress = [math]::Round(($currentSiteIndex / $siteCount) * 100, 0)
            $progress.UpdateProgress("ProcessSites", "Processing site $currentSiteIndex of $siteCount : $($siteItem.site)", $siteProgress)

            $siteResult = Invoke-SiteProcessing `
                -SiteItem $siteItem `
                -EPPData $eppDataResult.Data `
                -Environment $envConfig.Environment `
                -MonitoringConnection $dbResult.MonitoringConnection `
                -Logger $Logger `
                -ErrorTracker $ErrorTracker `
                -ResourceManager $ResourceManager `
                -Timestamp $Config.Timestamp `
                -Progress $progress `
                -SiteIndex $currentSiteIndex `
                -TotalSites $siteCount

            if ($siteResult.Success) {
                $sitesProcessed++

                # Aggregate all counts
                $totalRegularCollections += $siteResult.ProcessedCount
                $totalRetiredCollections += $siteResult.RetiredCollectionsProcessed
                $totalOrphanedCollections += $siteResult.OrphanedCollectionsProcessed
                $totalEppTaskSequences += $siteResult.TaskSequencesProcessed
                $totalRetiredTaskSequences += $siteResult.RetiredTaskSequencesProcessed
                $totalOrphanedTaskSequences += $siteResult.OrphanedTaskSequencesProcessed

                $Logger.Info("Site $($siteItem.site) totals - Regular: $($siteResult.ProcessedCount), " +
                    "Retired Coll: $($siteResult.RetiredCollectionsProcessed), " +
                    "Retired TS: $($siteResult.RetiredTaskSequencesProcessed), " +
                    "Orphaned Coll: $($siteResult.OrphanedCollectionsProcessed), " +
                    "Orphaned TS: $($siteResult.OrphanedTaskSequencesProcessed), " +
                    "EPP TS: $($siteResult.TaskSequencesProcessed)", "MainOrchestrator")
            }
            else {
                $sitesFailed++
            }
        }

        $progress.CompleteStep("ProcessSites")

        # Calculate grand totals
        $totalCollections = $totalRegularCollections + $totalRetiredCollections + $totalOrphanedCollections
        $totalTaskSequences = $totalEppTaskSequences + $totalRetiredTaskSequences + $totalOrphanedTaskSequences

        # Step 7: Finalize
        $progress.StartStep("Finalize", "Finalizing scan results")
        $timer.Stop()
        $totalMinutes = [math]::Round($timer.Elapsed.TotalMinutes, 2)

        $Logger.Info("=== SCCM Collection Scan Completed ===", "MainOrchestrator")
        $Logger.Info("Total processing time: $totalMinutes minutes", "MainOrchestrator")
        $Logger.Info("Sites processed: $sitesProcessed, Sites failed: $sitesFailed", "MainOrchestrator")
        $Logger.Info("=== Collection Breakdown ===", "MainOrchestrator")
        $Logger.Info("Regular Collections: $totalRegularCollections", "MainOrchestrator")
        $Logger.Info("Retired Collections: $totalRetiredCollections", "MainOrchestrator")
        $Logger.Info("Orphaned Collections: $totalOrphanedCollections", "MainOrchestrator")
        $Logger.Info("Total Collections: $totalCollections", "MainOrchestrator")
        $Logger.Info("=== Task Sequence Breakdown ===", "MainOrchestrator")
        $Logger.Info("EPP Cleanup Task Sequences: $totalEppTaskSequences", "MainOrchestrator")
        $Logger.Info("Retired Task Sequences: $totalRetiredTaskSequences", "MainOrchestrator")
        $Logger.Info("Orphaned Task Sequences: $totalOrphanedTaskSequences", "MainOrchestrator")
        $Logger.Info("Total Task Sequences: $totalTaskSequences", "MainOrchestrator")
        $Logger.Info("Error Summary: $($ErrorTracker.GetSummary())", "MainOrchestrator")

        $progress.CompleteStep("Finalize")
        $progress.Complete()

        return @{
            Success = -not $ErrorTracker.HasFatalErrors()
            TotalMinutes = $totalMinutes
            SitesProcessed = $sitesProcessed
            SitesFailed = $sitesFailed

            # Detailed collection counts
            RegularCollections = $totalRegularCollections
            RetiredCollections = $totalRetiredCollections
            OrphanedCollections = $totalOrphanedCollections
            TotalCollections = $totalCollections

            # Detailed task sequence counts
            EppTaskSequences = $totalEppTaskSequences
            RetiredTaskSequences = $totalRetiredTaskSequences
            OrphanedTaskSequences = $totalOrphanedTaskSequences
            TotalTaskSequences = $totalTaskSequences

            ErrorSummary = $ErrorTracker.GetSummary()
        }
    }
    catch {
        $timer.Stop()
        $progress.Complete()

        $ErrorTracker.RecordError("MainOrchestrator",
            "Fatal error in main execution",
            $true,
            $_.Exception)
        $Logger.Fatal("Main execution failed: $_", "MainOrchestrator")

        return @{
            Success = $false
            TotalMinutes = [math]::Round($timer.Elapsed.TotalMinutes, 2)
            SitesProcessed = 0
            SitesFailed = 0
            CollectionsProcessed = 0
            ErrorSummary = $ErrorTracker.GetSummary()
        }
    }
}
#endregion

#region Script Entry Point
# Main execution block
try {
    # Get script root
    $scriptRoot = Split-Path $MyInvocation.MyCommand.Path

    # Initialize configuration
    $config = [ScriptConfig]::new($scriptRoot, $environment, $days)
    $script:NonStagingDays = $days
    $script:KeepHistory = $KeepHistory.IsPresent
    # Round to millisecond precision (SQL Server datetime precision)
    $config.Timestamp = [datetime]::Parse($config.Timestamp.ToString("yyyy-MM-dd HH:mm:ss.fff"))

    # Initialize logger
    $logger = [Logger]::new($config.LogFile, $config.Environment)
    $logger.VerboseEnabled = ($VerbosePreference -eq 'Continue')

    # Initialize error tracker
    $errorTracker = [ErrorTracker]::new()

    # Initialize resource manager
    $resourceManager = [ResourceManager]::new($logger)

    # Import libraries at script level
    $logger.Info("Importing script libraries", "Main")
    $libraryResult = Import-ScriptLibraries `
        -ScriptPath $scriptRoot `
        -Logger $logger `
        -ErrorTracker $errorTracker

    if ($libraryResult.Success) {
        # Dot-source at script level so functions are globally available
        . $libraryResult.FunctionsPath
        $logger.Info("Functions loaded successfully from: $($libraryResult.FunctionsPath)", "Main")

        # Verify functions are available
        $adObjectCmd = Get-Command Get-ADObject -ErrorAction SilentlyContinue
        $adGroupCmd = Get-Command Get-ADGroup -ErrorAction SilentlyContinue

        if ($adObjectCmd) {
            $logger.Info("✓ Get-ADObject function is available", "Main")
            Write-Host "✓ Get-ADObject loaded" -ForegroundColor Green
        } else {
            $logger.Warning("✗ Get-ADObject function not found", "Main")
            Write-Host "✗ Get-ADObject NOT loaded" -ForegroundColor Red
        }

        if ($adGroupCmd) {
            $logger.Info("✓ Get-ADGroup function is available", "Main")
            Write-Host "✓ Get-ADGroup loaded" -ForegroundColor Green
        } else {
            $logger.Warning("✗ Get-ADGroup function not found", "Main")
            Write-Host "✗ Get-ADGroup NOT loaded" -ForegroundColor Red
        }
    } else {
        $logger.Warning("Failed to import libraries - AD functions may not be available", "Main")
    }

    # Log execution details
    $logger.Info("=== Script Execution Started ===", "Main")
    $logger.Info("Execution Timestamp: $($config.Timestamp.ToString('yyyy-MM-dd HH:mm:ss.fff'))", "Main")
    $logger.Info("Environment: $($config.Environment)", "Main")
    $logger.Info("Days Threshold: $($config.Days)", "Main")

    # Show log file summary
    Show-LogFileSummary -LogDir $config.LogDir

    # Execute main orchestrator
    $result = Invoke-SCCMCollectionScan `
        -Config $config `
        -Logger $logger `
        -ErrorTracker $errorTracker `
        -ResourceManager $resourceManager

    # Display results
    Write-Host "`n=== Execution Summary ===" -ForegroundColor Cyan
    Write-Host "Status: $(if ($result.Success) { 'SUCCESS' } else { 'FAILED' })" -ForegroundColor $(if ($result.Success) { 'Green' } else { 'Red' })
    Write-Host "Total Time: $($result.TotalMinutes) minutes" -ForegroundColor White
    Write-Host "Sites Processed: $($result.SitesProcessed)" -ForegroundColor White
    Write-Host "Sites Failed: $($result.SitesFailed)" -ForegroundColor White
    Write-Host "`n--- Collection Breakdown ---" -ForegroundColor Yellow
    Write-Host "Regular Collections: $($result.RegularCollections)" -ForegroundColor White
    Write-Host "Retired Collections: $($result.RetiredCollections)" -ForegroundColor White
    Write-Host "Orphaned Collections: $($result.OrphanedCollections)" -ForegroundColor White
    Write-Host "Total Collections: $($result.TotalCollections)" -ForegroundColor Cyan
    Write-Host "`n--- Task Sequence Breakdown ---" -ForegroundColor Yellow
    Write-Host "EPP Cleanup Task Sequences: $($result.EppTaskSequences)" -ForegroundColor White
    Write-Host "Retired Task Sequences: $($result.RetiredTaskSequences)" -ForegroundColor White
    Write-Host "Orphaned Task Sequences: $($result.OrphanedTaskSequences)" -ForegroundColor White
    Write-Host "Total Task Sequences: $($result.TotalTaskSequences)" -ForegroundColor Cyan
    Write-Host "`nErrors: $($result.ErrorSummary)" -ForegroundColor Yellow
    Write-Host "Log File: $($logger.GetCurrentLogFile())" -ForegroundColor White
    Write-Host "========================`n" -ForegroundColor Cyan

    # Export log analysis if there were errors
    if ($errorTracker.HasErrors()) {
        $analysisPath = Join-Path $config.LogDir "LogAnalysis_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').json"
        Export-LogAnalysis -LogFile $logger.GetCurrentLogFile() -OutputPath $analysisPath
    }

    # Set exit code
    if ($result.Success) {
        $exitCode = 0
    }
    elseif ($errorTracker.HasFatalErrors()) {
        $exitCode = 1
    }
    else {
        $exitCode = 2
    }
}
catch {
    Write-Host "`nCRITICAL ERROR: $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    $exitCode = 99
}
finally {
    # Cleanup resources
    if ($resourceManager) {
        Write-Host "Cleaning up resources..." -ForegroundColor Yellow
        $resourceManager.CleanupAll()
    }

    # Flush logger
    if ($logger) {
        $logger.Flush()
        Write-Host "`nLog file location: $($logger.GetCurrentLogFile())" -ForegroundColor Cyan
    }

    Write-Host "Script execution completed with exit code: $exitCode" -ForegroundColor $(
        switch ($exitCode) {
            0 { 'Green' }
            1 { 'Red' }
            2 { 'Yellow' }
            default { 'DarkRed' }
        }
    )
}

# Exit with appropriate code
exit $exitCode
#endregion