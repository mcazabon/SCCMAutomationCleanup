USE [MonitoringTestDB];
GO

-- EPP cleanup candidates with site name
CREATE OR ALTER VIEW dbo.vw_EPPCleanup
AS
SELECT  e.id,
        s.site                      AS SiteCode,
        s.environment               AS Environment,
        e.[Task Sequence Name]      AS TaskSequenceName,
        e.[Task Sequence Package ID] AS TaskSequencePackageID,
        e.[Task Sequence Path]      AS TaskSequencePath,
        e.[Task Sequence Status]    AS TaskSequenceStatus,
        e.[Task Sequence Message]   AS TaskSequenceMessage,
        e.[Referenced Package Name] AS ReferencedPackageName,
        e.[Reference Package ID]    AS ReferencePackageID,
        e.[EPP Status]              AS EPPStatus,
        e.[Match Type]              AS MatchType,
        e.[Collection ID]           AS CollectionID,
        e.[Collection Name]         AS CollectionName,
        e.[Collection Path]         AS CollectionPath,
        e.[Collection Type]         AS CollectionType,
        e.[Collection Status]       AS CollectionStatus,
        e.[Collection Message]      AS CollectionMessage,
        e.Actions,
        e.StatusTimestamp,
        e.[Timestamp]               AS ScanTimestamp
FROM dbo.tbl_SCCMTaskSequencesEPPCleanup e
LEFT JOIN dbo.tbl_SCCMSites s ON s.id = e.siteid;
GO

-- Retired-folder items that need review: active packages (Live, Pending Review)
-- or packages whose EPP status could not be verified (Not Found, Null, Error)
CREATE OR ALTER VIEW dbo.vw_EPPConflicts
AS
SELECT  SiteCode, Environment,
        CASE WHEN EPPStatus IN ('Live', 'Pending Review') THEN 'Active package'
             ELSE 'Unverified package' END AS ConflictType,
        ReferencedPackageName, ReferencePackageID, EPPStatus,
        TaskSequenceName, TaskSequencePackageID, TaskSequencePath, TaskSequenceStatus,
        CollectionID, CollectionName, CollectionPath, CollectionType, CollectionStatus,
        MatchType, Actions, ScanTimestamp
FROM dbo.vw_EPPCleanup
WHERE MatchType = 'Retired Folder'
  AND EPPStatus IN ('Live', 'Pending Review', 'Not Found', 'Null', 'Error');
GO

-- One row per scanned collection with status and referencing object counts
CREATE OR ALTER VIEW dbo.vw_CollectionObjectsScan
AS
SELECT  c.id,
        s.site                      AS SiteCode,
        s.environment               AS Environment,
        c.collectionsiteid          AS CollectionSiteID,
        c.collectionid              AS CollectionID,
        c.name                      AS CollectionName,
        c.path                      AS CollectionPath,
        c.collectiontype            AS CollectionType,
        c.deploymenttype            AS DeploymentType,
        c.membercount               AS MemberCount,
        c.lastmemberchangetime      AS LastMemberChangeTime,
        c.lastmemberchangedays      AS LastMemberChangeDays,
        c.LimitingCollection,
        c.ReferencedCollection,
        c.baselinecompliance        AS BaselineCompliance,
        c.baselineconfigitemscount  AS BaselineConfigItemsCount,
        c.baselinecollectionrefcnt  AS BaselineCollectionRefCount,
        c.compliancecnt             AS ComplianceCount,
        cs.status                   AS ScanStatus,
        cs.message                  AS ScanMessage,
        cs.actions                  AS ScanActions,
        ISNULL(ec.EPPCleanupCount, 0) AS EPPCleanupCount,
        ISNULL(tsc.TaskSequenceCount, 0) AS TaskSequenceCount,
        ISNULL(ap.ApplicationCount, 0) AS ApplicationCount,
        c.[timestamp]               AS ScanTimestamp
FROM dbo.tbl_SCCMCollections c
LEFT JOIN dbo.tbl_SCCMSites s ON s.id = c.siteid
LEFT JOIN dbo.tbl_SCCMCollectionStatus cs ON cs.collectionid = c.collectionid
OUTER APPLY (SELECT COUNT(*) AS EPPCleanupCount FROM dbo.tbl_SCCMTaskSequencesEPPCleanup x
             WHERE x.[Collection ID] = c.collectionsiteid) ec
OUTER APPLY (SELECT COUNT(*) AS TaskSequenceCount FROM dbo.tbl_SCCMtasksequences t
             WHERE t.collectionid = c.collectionid) tsc
OUTER APPLY (SELECT COUNT(*) AS ApplicationCount FROM dbo.tbl_SCCMapplications a
             WHERE a.collectionid = c.collectionid) ap;
GO
