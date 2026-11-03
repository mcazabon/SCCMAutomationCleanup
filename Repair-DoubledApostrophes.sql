/*
    One-time repair for monitoring rows saved before the scan stopped doubling apostrophes
    (O''Brien -> O'Brien). Safe to re-run. Review the counts, then COMMIT or ROLLBACK.
    Not needed for sites that have been re-scanned without -KeepHistory.
*/
USE [MonitoringTestDB];
GO

SET XACT_ABORT ON;
BEGIN TRAN;

DECLARE @q nvarchar(2) = N'''''';

UPDATE dbo.tbl_SCCMCollections
SET name = REPLACE(name, @q, N''''), path = REPLACE(path, @q, N''''),
    LimitingCollection = REPLACE(LimitingCollection, @q, N'''')
WHERE name LIKE N'%''''%' OR path LIKE N'%''''%' OR LimitingCollection LIKE N'%''''%';
PRINT 'tbl_SCCMCollections: ' + CAST(@@ROWCOUNT AS varchar(20));

UPDATE dbo.tbl_SCCMapplications
SET name = REPLACE(name, @q, N''''), path = REPLACE(path, @q, N'''')
WHERE name LIKE N'%''''%' OR path LIKE N'%''''%';
PRINT 'tbl_SCCMapplications: ' + CAST(@@ROWCOUNT AS varchar(20));

UPDATE dbo.tbl_SCCMtasksequences
SET name = REPLACE(name, @q, N'''')
WHERE name LIKE N'%''''%';
PRINT 'tbl_SCCMtasksequences: ' + CAST(@@ROWCOUNT AS varchar(20));

UPDATE dbo.tbl_SCCMTaskSequencesEPPCleanup
SET [Task Sequence Name]     = REPLACE([Task Sequence Name], @q, N''''),
    [Task Sequence Path]     = REPLACE([Task Sequence Path], @q, N''''),
    [Referenced Package Name] = REPLACE([Referenced Package Name], @q, N''''),
    [Collection Name]        = REPLACE([Collection Name], @q, N''''),
    [Collection Path]        = REPLACE([Collection Path], @q, N'''')
WHERE [Task Sequence Name] LIKE N'%''''%' OR [Task Sequence Path] LIKE N'%''''%'
   OR [Referenced Package Name] LIKE N'%''''%' OR [Collection Name] LIKE N'%''''%'
   OR [Collection Path] LIKE N'%''''%';
PRINT 'tbl_SCCMTaskSequencesEPPCleanup: ' + CAST(@@ROWCOUNT AS varchar(20));

COMMIT;
GO
