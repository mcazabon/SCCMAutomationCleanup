/*
    SCCM Collection Cleanup - monitoring database schema
    Target instance : <your SQL Server instance>
    Target database : <your monitoring database - see USE below>

    Derived from the INSERT/SELECT statements in SCCMCollectionScan.ps1.
    Plain T-SQL (no SQLCMD mode needed) - runs directly in SSMS.
    Idempotent: safe to re-run; existing tables and data are left untouched.
*/
IF DB_ID(N'MonitoringTestDB') IS NULL
    CREATE DATABASE [MonitoringTestDB];
GO

USE [MonitoringTestDB];
GO

/* ------------------------------------------------------------------ */
/* tbl_SCCMSites - site registration (Register-SCCMSite)               */
/* ------------------------------------------------------------------ */
IF OBJECT_ID(N'dbo.tbl_SCCMSites', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.tbl_SCCMSites
    (
        [id]          INT IDENTITY(1,1) NOT NULL,
        [site]        NVARCHAR(50)      NOT NULL,
        [environment] NVARCHAR(50)      NOT NULL,
        [timestamp]   DATETIME          NOT NULL CONSTRAINT DF_tbl_SCCMSites_timestamp DEFAULT (GETDATE()),
        CONSTRAINT PK_tbl_SCCMSites PRIMARY KEY CLUSTERED ([id]),
        CONSTRAINT UQ_tbl_SCCMSites_site_environment UNIQUE ([site], [environment])
    );
END
GO

/* ------------------------------------------------------------------ */
/* tbl_SCCMCollections - collection cleanup candidates                 */
/* ------------------------------------------------------------------ */
IF OBJECT_ID(N'dbo.tbl_SCCMCollections', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.tbl_SCCMCollections
    (
        [id]                          INT IDENTITY(1,1) NOT NULL,
        [siteid]                      INT               NOT NULL,
        [collectionid]                NVARCHAR(50)      NOT NULL,
        [collectionsiteid]            NVARCHAR(50)      NULL,
        [name]                        NVARCHAR(500)     NOT NULL,
        [path]                        NVARCHAR(500)     NULL,
        [lastmemberchangetime]        DATETIME          NULL,
        [lastmemberchangedays]        INT               NULL,
        [adaptivaclientsettingsref]   BIT               NOT NULL CONSTRAINT DF_tbl_SCCMCollections_acs DEFAULT (0),
        [adaptivapolicyref]           BIT               NOT NULL CONSTRAINT DF_tbl_SCCMCollections_apr DEFAULT (0),
        [adaptivahealthcheckref]      BIT               NOT NULL CONSTRAINT DF_tbl_SCCMCollections_ahr DEFAULT (0),
        [membercount]                 INT               NULL,
        [collectiontype]              NVARCHAR(100)     NULL,
        [deploymenttype]              NVARCHAR(100)     NULL,
        [baselinetotaldevicecnt]      NVARCHAR(50)      NULL,
        [compliancecnt]               NVARCHAR(50)      NULL,
        [baselinecompliance]          NVARCHAR(50)      NULL,
        [baselinecollectionrefcnt]    NVARCHAR(50)      NULL,
        [baselineconfigitemscount]    NVARCHAR(50)      NULL,
        [baselineconfigitemreference] NVARCHAR(50)      NULL,
        [baselinecollectionreference] NVARCHAR(50)      NULL,
        [LimitingCollection]          NVARCHAR(500)     NULL,
        [ReferencedCollection]        NVARCHAR(MAX)     NULL,
        [timestamp]                   DATETIME          NOT NULL CONSTRAINT DF_tbl_SCCMCollections_timestamp DEFAULT (GETDATE()),
        CONSTRAINT PK_tbl_SCCMCollections PRIMARY KEY CLUSTERED ([id]),
        CONSTRAINT UQ_tbl_SCCMCollections_collectionid UNIQUE ([collectionid]),
        CONSTRAINT FK_tbl_SCCMCollections_siteid FOREIGN KEY ([siteid]) REFERENCES dbo.tbl_SCCMSites ([id])
    );

    CREATE NONCLUSTERED INDEX IX_tbl_SCCMCollections_siteid ON dbo.tbl_SCCMCollections ([siteid]);
END
GO

/* ------------------------------------------------------------------ */
/* tbl_SCCMapplications - application / package references             */
/* ------------------------------------------------------------------ */
/* [appref] is written by Add-ApplicationToDatabase; [apppref] is the  */
/* spelling used by Add-RetiredTaskSequenceToDatabase. Both are kept   */
/* so every INSERT in the script works unchanged.                      */
IF OBJECT_ID(N'dbo.tbl_SCCMapplications', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.tbl_SCCMapplications
    (
        [id]               INT IDENTITY(1,1) NOT NULL,
        [collectionid]     NVARCHAR(50)      NULL,
        [collectionsiteid] NVARCHAR(50)      NULL,
        [name]             NVARCHAR(500)     NULL,
        [path]             NVARCHAR(500)     NULL,
        [packageid]        NVARCHAR(50)      NULL,
        [packagesizeKB]    NVARCHAR(50)      NULL,
        [packagetype]      NVARCHAR(100)     NULL,
        [App_Comments]     NVARCHAR(MAX)     NULL,
        [appref]           BIT               NULL,
        [apppref]          BIT               NULL,
        [eppstatus]        NVARCHAR(100)     NULL,
        [timestamp]        DATETIME          NOT NULL CONSTRAINT DF_tbl_SCCMapplications_timestamp DEFAULT (GETDATE()),
        CONSTRAINT PK_tbl_SCCMapplications PRIMARY KEY CLUSTERED ([id])
    );

    CREATE NONCLUSTERED INDEX IX_tbl_SCCMapplications_packageid ON dbo.tbl_SCCMapplications ([packageid]);
    CREATE NONCLUSTERED INDEX IX_tbl_SCCMapplications_collectionid ON dbo.tbl_SCCMapplications ([collectionid]);
END
GO

/* ------------------------------------------------------------------ */
/* tbl_SCCMtasksequences - task sequence references                    */
/* ------------------------------------------------------------------ */
/* collectionid holds NULL (orphaned) or '-1' (retired folder) too.    */
IF OBJECT_ID(N'dbo.tbl_SCCMtasksequences', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.tbl_SCCMtasksequences
    (
        [id]               INT IDENTITY(1,1) NOT NULL,
        [collectionid]     NVARCHAR(50)      NULL,
        [collectionsiteid] NVARCHAR(50)      NULL,
        [name]             NVARCHAR(500)     NULL,
        [apptsref]         BIT               NULL,
        [timestamp]        DATETIME          NOT NULL CONSTRAINT DF_tbl_SCCMtasksequences_timestamp DEFAULT (GETDATE()),
        CONSTRAINT PK_tbl_SCCMtasksequences PRIMARY KEY CLUSTERED ([id])
    );

    CREATE NONCLUSTERED INDEX IX_tbl_SCCMtasksequences_collectionid ON dbo.tbl_SCCMtasksequences ([collectionid]);
    CREATE NONCLUSTERED INDEX IX_tbl_SCCMtasksequences_collectionsiteid ON dbo.tbl_SCCMtasksequences ([collectionsiteid]);
END
GO

/* ------------------------------------------------------------------ */
/* tbl_SCCMCollectionStatus - processing status / recommended actions  */
/* ------------------------------------------------------------------ */
/* collectionid also receives package IDs for orphaned task sequences. */
IF OBJECT_ID(N'dbo.tbl_SCCMCollectionStatus', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.tbl_SCCMCollectionStatus
    (
        [id]           INT IDENTITY(1,1) NOT NULL,
        [collectionid] NVARCHAR(50)      NOT NULL,
        [status]       NVARCHAR(50)      NOT NULL,
        [message]      NVARCHAR(MAX)     NULL,
        [actions]      NVARCHAR(500)     NULL,
        [timestamp]    DATETIME          NOT NULL CONSTRAINT DF_tbl_SCCMCollectionStatus_timestamp DEFAULT (GETDATE()),
        CONSTRAINT PK_tbl_SCCMCollectionStatus PRIMARY KEY CLUSTERED ([id]),
        CONSTRAINT UQ_tbl_SCCMCollectionStatus_collectionid UNIQUE ([collectionid]),
        CONSTRAINT CK_tbl_SCCMCollectionStatus_status
            CHECK ([status] IN (N'Discovered', N'Archived', N'Review Required', N'Orphaned', N'None'))
    );
END
GO

/* ------------------------------------------------------------------ */
/* tbl_SCCMTaskSequencesEPPCleanup - EPP cleanup candidates            */
/* ------------------------------------------------------------------ */
IF OBJECT_ID(N'dbo.tbl_SCCMTaskSequencesEPPCleanup', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.tbl_SCCMTaskSequencesEPPCleanup
    (
        [id]                        INT IDENTITY(1,1) NOT NULL,
        [siteid]                    INT               NOT NULL,
        [Task Sequence Name]        NVARCHAR(500)     NULL,
        [Task Sequence Package ID]  NVARCHAR(50)      NULL,
        [Task Sequence Path]        NVARCHAR(500)     NULL,
        [Referenced Package Name]   NVARCHAR(500)     NULL,
        [Reference Package ID]      NVARCHAR(50)      NULL,
        [Collection Name]           NVARCHAR(500)     NULL,
        [Collection ID]             NVARCHAR(50)      NULL,
        [Collection Type]           NVARCHAR(100)     NULL,
        [Collection Path]           NVARCHAR(500)     NULL,
        [EPP Status]                NVARCHAR(100)     NULL,
        [Match Type]                NVARCHAR(100)     NULL,
        [Timestamp]                 DATETIME          NOT NULL CONSTRAINT DF_tbl_SCCMTSEPP_timestamp DEFAULT (GETDATE()),
        [Task Sequence Message]     NVARCHAR(MAX)     NULL,
        [Collection Message]        NVARCHAR(MAX)     NULL,
        [Task Sequence Status]      NVARCHAR(50)      NULL,
        [Collection Status]         NVARCHAR(50)      NULL,
        [Actions]                   NVARCHAR(500)     NULL,
        [StatusTimestamp]           DATETIME          NULL,
        CONSTRAINT PK_tbl_SCCMTaskSequencesEPPCleanup PRIMARY KEY CLUSTERED ([id]),
        CONSTRAINT FK_tbl_SCCMTSEPP_siteid FOREIGN KEY ([siteid]) REFERENCES dbo.tbl_SCCMSites ([id]),
        CONSTRAINT CK_tbl_SCCMTSEPP_tsstatus
            CHECK ([Task Sequence Status] IS NULL
                   OR [Task Sequence Status] IN (N'Discovered', N'Archived', N'Review Required', N'Orphaned', N'None')),
        CONSTRAINT CK_tbl_SCCMTSEPP_collstatus
            CHECK ([Collection Status] IS NULL
                   OR [Collection Status] IN (N'Discovered', N'Archived', N'Review Required', N'Orphaned', N'None'))
    );

    CREATE NONCLUSTERED INDEX IX_tbl_SCCMTSEPP_site_pkg
        ON dbo.tbl_SCCMTaskSequencesEPPCleanup ([siteid], [Task Sequence Package ID]);
END
GO

PRINT 'Monitoring schema deployed.';
GO
