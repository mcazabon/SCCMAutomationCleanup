/*
    Compatibility objects for a MOCK SCCM database (test/UAT only).

    The real SCCM site database exposes the views/functions below; the mock stores the same data
    under other names (or not at all). These objects map the mock tables so the unmodified
    SCCMCollectionScan.ps1 queries (task sequence, orphan and collection detection) can run.
    Data that the mock does not contain (application relations, task sequence application
    references) is exposed as empty sets.

    Do NOT deploy to a real site database. Idempotent: safe to re-run.
*/
USE [SCCM];
GO

/* ---------- Folders ---------- */
CREATE OR ALTER VIEW dbo.Folders
AS
SELECT ContainerNodeID, Name, ISNULL(ParentContainerNodeID, 0) AS ParentContainerNodeID, ObjectType, FolderPath
FROM dbo.ObjectContainerNode;
GO

CREATE OR ALTER VIEW dbo.vSMS_Folders
AS
SELECT ContainerNodeID, Name, ISNULL(ParentContainerNodeID, 0) AS ParentContainerNodeID, ObjectType, FolderPath
FROM dbo.ObjectContainerNode;
GO

CREATE OR ALTER VIEW dbo.vFolderMembers
AS
SELECT InstanceKey, ObjectType, ContainerNodeID
FROM dbo.ObjectContainerItem;
GO

CREATE OR ALTER VIEW dbo.FolderMembers
AS
SELECT InstanceKey, ObjectType, ContainerNodeID
FROM dbo.ObjectContainerItem;
GO

/* ---------- Existing mock views extended with the columns the queries use ---------- */
CREATE OR ALTER VIEW dbo.v_ConfigurationItems
AS
SELECT CI_ID, CI_UniqueID, CIType_ID, IsExpired, IsEnabled, CIVersion, DateCreated, DateLastModified, CreatedBy,
       CI_UniqueID AS ModelID
FROM dbo.CI_ConfigurationItems;
GO

CREATE OR ALTER VIEW dbo.v_Applications
AS
SELECT CI_ID, CI_UniqueID, DisplayName, Description, IsExpired, IsEnabled, CIVersion, DateCreated, DateLastModified, CreatedBy,
       Description AS AdminComments
FROM dbo.CI_ConfigurationItems WHERE CIType_ID = 10;
GO

CREATE OR ALTER VIEW dbo.v_DeploymentSummary
AS
SELECT d.DeploymentID, d.SoftwareName, d.FeatureType, d.CollectionID, d.PackageID, d.CI_ID, d.DeploymentIntent,
       d.DeploymentTime, d.EnforcementDeadline, d.NumberTargeted, d.NumberSuccess, d.NumberInProgress,
       d.NumberErrors, d.NumberUnknown,
       (SELECT c.CI_UniqueID FROM dbo.CI_ConfigurationItems c WHERE c.CI_ID = d.CI_ID) AS ModelID
FROM dbo.DeploymentSummary d;
GO

CREATE OR ALTER VIEW dbo.v_Package
AS
SELECT PackageID, Name, Version, Manufacturer, PackageType, SourceSize, SourceDate, LastRefreshTime,
       CAST(NULL AS NVARCHAR(255)) AS Description,
       CAST(NULL AS NVARCHAR(50))  AS Language,
       CAST(N'MOCK' AS NVARCHAR(10)) AS SourceSite
FROM dbo.SoftwareList;
GO

CREATE OR ALTER VIEW dbo.v_TaskSequencePackage
AS
SELECT PackageID, Name, Version, Manufacturer, SourceDate, LastRefreshTime,
       CAST(NULL AS NVARCHAR(255)) AS Description
FROM dbo.SoftwareList WHERE PackageType = 4;
GO

CREATE OR ALTER VIEW dbo.v_PackageStatusDistPointsSumm
AS
SELECT PackageID, NALPath, Status, LastUpdateTime, NALPath AS ServerNALPath, Status AS InstallStatus
FROM dbo.PkgStatus;
GO

/* ---------- Task sequences / deployments ---------- */
CREATE OR ALTER VIEW dbo.v_TaskSequenceReferencesInfo
AS
SELECT PackageID, ReferencedPackageID AS ReferencePackageID
FROM dbo.TaskSequenceSteps
WHERE ReferencedPackageID IS NOT NULL;
GO

CREATE OR ALTER VIEW dbo.v_TaskSequencePackageReferences
AS
SELECT PackageID, ReferencedPackageID AS RefPackageID
FROM dbo.TaskSequenceSteps
WHERE ReferencedPackageID IS NOT NULL;
GO

CREATE OR ALTER VIEW dbo.v_TaskSequenceAppReferencesInfo
AS
SELECT CAST(NULL AS NVARCHAR(50)) AS PackageID, CAST(NULL AS NVARCHAR(256)) AS RefAppModelName
WHERE 1 = 0;
GO

CREATE OR ALTER VIEW dbo.v_Advertisement
AS
SELECT DeploymentID AS AdvertisementID, PackageID, CollectionID,
       SoftwareName AS AdvertisementName,
       CAST(CASE WHEN DeploymentIntent = 1 THEN 0x20 ELSE 0 END AS INT) AS AdvertFlags
FROM dbo.DeploymentSummary
WHERE PackageID IS NOT NULL;
GO

CREATE OR ALTER VIEW dbo.v_CIAssignment
AS
SELECT DeploymentID AS AssignmentID, CollectionID
FROM dbo.DeploymentSummary
WHERE FeatureType IN (1, 6);
GO

CREATE OR ALTER VIEW dbo.v_CIAssignmentToCI
AS
SELECT DeploymentID AS AssignmentID, CI_ID
FROM dbo.DeploymentSummary
WHERE FeatureType IN (1, 6) AND CI_ID IS NOT NULL;
GO

CREATE OR ALTER FUNCTION dbo.fn_DeploymentSummary (@LocaleID INT)
RETURNS TABLE
AS
RETURN
    SELECT s.SoftwareName, c.CollectionName, s.FeatureType, s.CollectionID
    FROM dbo.DeploymentSummary s
    INNER JOIN dbo.Collections_G c ON c.CollectionID = s.CollectionID;
GO

/* ---------- Collections ---------- */
CREATE OR ALTER VIEW dbo.v_Collections
AS
SELECT TRY_CAST(SUBSTRING(CollectionID, 4, 20) AS INT) AS CollectionID,
       CollectionID AS SiteID,
       CollectionName,
       MemberCount,
       LastChangeTime AS LastMemberChangeTime,
       CollectionType,
       LimitToCollectionID,
       (SELECT l.CollectionName FROM dbo.Collections_G l WHERE l.CollectionID = c.LimitToCollectionID) AS LimitToCollectionName,
       CAST(0 AS INT) AS PowerConfigsCount
FROM dbo.Collections_G c;
GO

CREATE OR ALTER VIEW dbo.Collection_Rules
AS
SELECT CollectionID, ReferencedCollectionID
FROM dbo.CollectionIncludeExclude;
GO

CREATE OR ALTER VIEW dbo.vSMS_ClientSettingsAssignments
AS
SELECT ClientSettingsID, CollectionID
FROM dbo.ClientSettingsAssignments;
GO

CREATE OR ALTER VIEW dbo.v_FullCollectionMembership_Valid
AS
SELECT CollectionID, ResourceID, Name, SiteCode, IsClient, IsActive, IsDirect
FROM dbo.v_FullCollectionMembership;
GO

/* ---------- Applications ---------- */
CREATE OR ALTER FUNCTION dbo.fn_ListApplicationCIs (@LocaleID INT)
RETURNS TABLE
AS
RETURN
    SELECT a.CI_ID,
           a.CI_UniqueID AS ModelName,
           a.DisplayName,
           a.CreatedBy,
           a.IsEnabled,
           a.CIVersion,
           CAST(1 AS BIT) AS IsLatest,
           CAST(0 AS BIT) AS IsSuperseded,
           CAST(CASE WHEN EXISTS (SELECT 1 FROM dbo.DeploymentSummary s WHERE s.CI_ID = a.CI_ID AND s.FeatureType = 1) THEN 1 ELSE 0 END AS BIT) AS IsDeployed,
           (SELECT COUNT(*) FROM dbo.ApplicationDeploymentTypes d WHERE d.CI_ID = a.CI_ID) AS NumberOfDeploymentTypes,
           (SELECT COUNT(*) FROM dbo.DeploymentSummary s WHERE s.CI_ID = a.CI_ID AND s.FeatureType = 1) AS NumberOfDeployments,
           CAST(0 AS INT) AS NumberOfDevicesWithApp,
           CAST(0 AS INT) AS NumberOfDevicesWithFailure,
           (SELECT TOP 1 n.FolderPath
            FROM dbo.ObjectContainerItem i
            INNER JOIN dbo.ObjectContainerNode n ON n.ContainerNodeID = i.ContainerNodeID
            WHERE i.ObjectType = 6000 AND i.InstanceKey = a.CI_UniqueID) AS ObjectPath
    FROM dbo.CI_ConfigurationItems a
    WHERE a.CIType_ID = 10;
GO

CREATE OR ALTER FUNCTION dbo.fn_ListApplicationCIs_List (@LocaleID INT)
RETURNS TABLE
AS
RETURN SELECT * FROM dbo.fn_ListApplicationCIs(@LocaleID);
GO

CREATE OR ALTER FUNCTION dbo.fn_ListLatestApplicationCIs (@LocaleID INT)
RETURNS TABLE
AS
RETURN SELECT * FROM dbo.fn_ListApplicationCIs(@LocaleID);
GO

CREATE OR ALTER FUNCTION dbo.fn_ListDeploymentTypeCIs (@LocaleID INT)
RETURNS TABLE
AS
RETURN
    SELECT a.CI_UniqueID AS AppModelName,
           d.Technology,
           CAST(1 AS BIT) AS IsLatest,
           d.DeploymentTypeID,
           d.DeploymentTypeName AS DisplayName
    FROM dbo.ApplicationDeploymentTypes d
    INNER JOIN dbo.CI_ConfigurationItems a ON a.CI_ID = d.CI_ID;
GO

CREATE OR ALTER VIEW dbo.vSMS_AppRelation_Flat
AS
SELECT CAST(NULL AS INT) AS FromApplicationCIID, CAST(NULL AS INT) AS ToApplicationCIID
WHERE 1 = 0;
GO

/* ---------- Baselines / configuration items ---------- */
CREATE OR ALTER VIEW dbo.v_CITypes
AS
SELECT CIType_ID, TypeName
FROM (VALUES (2, N'Baseline'), (3, N'Software Update'), (10, N'Application')) AS t (CIType_ID, TypeName);
GO

CREATE OR ALTER VIEW dbo.v_CIRelation
AS
SELECT BaselineCI_ID AS FromCIID, ConfigItemCI_ID AS ToCIID
FROM dbo.BaselineCIRelation;
GO

CREATE OR ALTER FUNCTION dbo.fn_ListCI_ComplianceState (@LocaleID INT)
RETURNS TABLE
AS
RETURN
    SELECT b.MachineID AS ResourceID,
           b.BaselineCI_ID AS CI_ID,
           ci.CIType_ID,
           CASE WHEN b.ComplianceState = 1 THEN N'Compliant' ELSE N'Non-Compliant' END AS ComplianceStateName
    FROM dbo.BaselineComplianceStatus b
    INNER JOIN dbo.CI_ConfigurationItems ci ON ci.CI_ID = b.BaselineCI_ID;
GO

/* ---------- Distribution points / package status ---------- */
CREATE OR ALTER VIEW dbo.v_DistributionPoint
AS
SELECT PackageID, NALPath AS ServerNALPath, CAST(0 AS INT) AS IsPeerDP, LastUpdateTime AS LastRefreshTime
FROM dbo.PkgStatus;
GO

CREATE OR ALTER VIEW dbo.v_PackageStatus
AS
SELECT PackageID, NALPath AS PkgServer, Status
FROM dbo.PkgStatus;
GO

CREATE OR ALTER VIEW dbo.v_PackageStatusRootSummarizer
AS
SELECT PackageID, SourceSize, SourceSize AS SourceCompressedSize
FROM dbo.SoftwareList;
GO

CREATE OR ALTER VIEW dbo.vSMS_DpGroupInfo
AS
SELECT CAST(N'DPG00001' AS NVARCHAR(20)) AS GroupID, CAST(N'All DPs' AS NVARCHAR(20)) AS Name;
GO

CREATE OR ALTER VIEW dbo.v_dpgroupPackages
AS
SELECT CAST(N'DPG00001' AS NVARCHAR(20)) AS GroupID, PackageID AS PkgID
FROM dbo.SoftwareList;
GO

CREATE OR ALTER VIEW dbo.v_dpgroupcontentdetails
AS
SELECT CAST(N'DPG00001' AS NVARCHAR(20)) AS GroupID,
       PackageID AS PkgID,
       COUNT(*) AS TargetedDPCount,
       SUM(CASE WHEN Status = 0 THEN 1 ELSE 0 END) AS NumberInstalled,
       CAST(0 AS INT) AS NumberInProgress,
       SUM(CASE WHEN Status = 2 THEN 1 ELSE 0 END) AS NumberErrors
FROM dbo.PkgStatus
GROUP BY PackageID;
GO
