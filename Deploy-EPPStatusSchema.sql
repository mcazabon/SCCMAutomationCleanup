/*
    EPP status - software package/application staging
    Target instance : <your SQL Server instance>
    Target database : SPDW_Staging

    dbo.spdw_SFSoftwarePackages holds one row per SCCM package or application,
    so every item has an EPP status record. dbo.usp_SyncSFSoftwarePackages
    refreshes the list from the primary site database (default: SCCM) and never
    overwrites an EPPStatus that has already been set.
    dbo.spdw_vwSFSoftwarePackages is read by SCCMCollectionScan.ps1 (EPPQuery).
    Idempotent: safe to re-run.
*/
IF DB_ID(N'SPDW_Staging') IS NULL
    CREATE DATABASE [SPDW_Staging];
GO

USE [SPDW_Staging];
GO

IF OBJECT_ID(N'dbo.spdw_SFSoftwarePackages', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.spdw_SFSoftwarePackages
    (
        [id]               INT IDENTITY(1,1) NOT NULL,
        [SourceType]       NVARCHAR(20)      NOT NULL,   -- Package | Application
        [SourceID]         NVARCHAR(256)     NOT NULL,   -- PackageID | CI_UniqueID
        [PackageName]      NVARCHAR(500)     NOT NULL,
        [PackageShortName] NVARCHAR(255)     NULL,       -- app AdminComments / package Comment (matched by the script)
        [EPPStatus]        NVARCHAR(100)     NULL,       -- set from the EPP source of truth
        [InSource]         BIT               NOT NULL CONSTRAINT DF_spdw_SFSoftwarePackages_InSource DEFAULT (1),
        [FirstSeen]        DATETIME          NOT NULL CONSTRAINT DF_spdw_SFSoftwarePackages_FirstSeen DEFAULT (GETDATE()),
        [LastSynced]       DATETIME          NOT NULL CONSTRAINT DF_spdw_SFSoftwarePackages_LastSynced DEFAULT (GETDATE()),
        CONSTRAINT PK_spdw_SFSoftwarePackages PRIMARY KEY CLUSTERED ([id]),
        CONSTRAINT UQ_spdw_SFSoftwarePackages_Source UNIQUE ([SourceType], [SourceID])
    );

    CREATE NONCLUSTERED INDEX IX_spdw_SFSoftwarePackages_PackageName
        ON dbo.spdw_SFSoftwarePackages ([PackageName]);
END
GO

/* Upgrade a table created by an earlier version of this script */
IF COL_LENGTH(N'dbo.spdw_SFSoftwarePackages', N'SourceType') IS NULL
BEGIN
    ALTER TABLE dbo.spdw_SFSoftwarePackages ADD
        [SourceType] NVARCHAR(20)  NOT NULL CONSTRAINT DF_spdw_SFSoftwarePackages_SourceType DEFAULT (N'Package'),
        [SourceID]   NVARCHAR(256) NOT NULL CONSTRAINT DF_spdw_SFSoftwarePackages_SourceID DEFAULT (N''),
        [InSource]   BIT           NOT NULL CONSTRAINT DF_spdw_SFSoftwarePackages_InSource DEFAULT (1),
        [FirstSeen]  DATETIME      NOT NULL CONSTRAINT DF_spdw_SFSoftwarePackages_FirstSeen DEFAULT (GETDATE()),
        [LastSynced] DATETIME      NOT NULL CONSTRAINT DF_spdw_SFSoftwarePackages_LastSynced DEFAULT (GETDATE());
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'UQ_spdw_SFSoftwarePackages_Source' AND object_id = OBJECT_ID(N'dbo.spdw_SFSoftwarePackages'))
    EXEC (N'UPDATE dbo.spdw_SFSoftwarePackages SET SourceID = CONCAT(N''legacy-'', id) WHERE SourceID = N'''';
            ALTER TABLE dbo.spdw_SFSoftwarePackages ADD CONSTRAINT UQ_spdw_SFSoftwarePackages_Source UNIQUE ([SourceType], [SourceID]);');
GO

CREATE OR ALTER VIEW dbo.spdw_vwSFSoftwarePackages
AS
SELECT
    [EPPStatus]        AS [EPP Status],
    [PackageName]      AS [Package Name],
    [PackageShortName] AS [Package Short Name]
FROM dbo.spdw_SFSoftwarePackages;
GO

CREATE OR ALTER PROCEDURE dbo.usp_SyncSFSoftwarePackages
    @SccmDb SYSNAME = N'SCCM'
AS
BEGIN
    SET NOCOUNT ON;

    IF DB_ID(@SccmDb) IS NULL
    BEGIN
        RAISERROR(N'SCCM database %s not found on this instance.', 16, 1, @SccmDb);
        RETURN;
    END

    DECLARE @db NVARCHAR(300) = QUOTENAME(@SccmDb);
    DECLARE @sql NVARCHAR(MAX) = N'
    SELECT N''Package'' AS SourceType, CAST(p.PackageID AS NVARCHAR(256)) AS SourceID,
           CAST(p.Name AS NVARCHAR(500)) AS PackageName, CAST(NULL AS NVARCHAR(255)) AS PackageShortName
    FROM ' + @db + N'.dbo.v_Package p WITH (NOLOCK)
    WHERE p.PackageType = 0
    UNION ALL
    SELECT N''Application'', CAST(a.CI_UniqueID AS NVARCHAR(256)),
           CAST(a.DisplayName AS NVARCHAR(500)), CAST(a.AdminComments AS NVARCHAR(255))
    FROM ' + @db + N'.dbo.fn_ListApplicationCIs(1033) a
    WHERE a.IsLatest = 1;';

    CREATE TABLE #src
    (
        SourceType       NVARCHAR(20)  NOT NULL,
        SourceID         NVARCHAR(256) NOT NULL,
        PackageName      NVARCHAR(500) NOT NULL,
        PackageShortName NVARCHAR(255) NULL
    );
    INSERT INTO #src EXEC (@sql);

    MERGE dbo.spdw_SFSoftwarePackages AS t
    USING (SELECT DISTINCT SourceType, SourceID, PackageName, PackageShortName FROM #src) AS s
        ON t.SourceType = s.SourceType AND t.SourceID = s.SourceID
    WHEN MATCHED THEN
        UPDATE SET PackageName = s.PackageName, PackageShortName = s.PackageShortName,
                   InSource = 1, LastSynced = GETDATE()
    WHEN NOT MATCHED BY TARGET THEN
        INSERT (SourceType, SourceID, PackageName, PackageShortName)
        VALUES (s.SourceType, s.SourceID, s.PackageName, s.PackageShortName)
    WHEN NOT MATCHED BY SOURCE AND t.InSource = 1 THEN
        UPDATE SET InSource = 0, LastSynced = GETDATE();
END
GO
