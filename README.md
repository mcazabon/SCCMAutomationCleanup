# SCCM Cleanup Automation

`SCCMCollectionScan.ps1` scans SCCM collections, task sequences and applications for cleanup candidates, cross-checks them against EPP status data, and writes results to a monitoring database.

## Setup

1. Copy `Library\Resources.sample.xml` to `Library\Resources.xml` (git-ignored) and fill in your servers and database names.
2. Deploy the SQL: `Deploy-EPPStatusSchema.sql` (staging), `Deploy-MonitoringSchema.sql`, then `Deploy-MonitoringViews.sql`.
3. Optional, SQL authentication: set `SCCM_SQL_AUTH_SERVERS` (regex of server names) and `SCCM_SQL_SA_PASSWORD`. Other servers use Windows authentication.
4. `Deploy-SCCMCompatViews.sql` is only for a mock SCCM test database. Never run it against a real site database.

## Run

```powershell
.\SCCMCollectionScan.ps1 -environment uat -days 365
```

See `Get-Help .\SCCMCollectionScan.ps1 -Full` for parameters.

## Tests

```powershell
.\Tests\Test-ProcessingRules.ps1   # unit tests, no database needed
.\Tests\Test-Smoke.ps1 [-RunScan]  # pre-flight checks against your configured databases
```
