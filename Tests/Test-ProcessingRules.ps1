. "$PSScriptRoot\..\Library\Functions.ps1"

$logger = [PSCustomObject]@{}
foreach ($m in 'Debug','Info','Warning','Error','Fatal') { $logger | Add-Member -MemberType ScriptMethod -Name $m -Value { param($a, $b) } }

$ref = Get-Date
$noAd = @{ Exists = $false; WhenChanged = $null }

function New-Item2($path, $days, $extra = @{}) {
    $o = [PSCustomObject]@{
        CollectionPath = $path; 'Days Since Last Member Change' = $days
        LimitingCollection = 'False'; ReferencedCollection = 'False'; 'App TS References' = 'False'
        'Adaptiva Client Settings Collection Reference' = $false
        'Adaptiva Policy Collection Reference' = $false
        'Adaptiva Health Check Collection Reference' = $false
    }
    foreach ($k in $extra.Keys) { $o.$k = $extra[$k] }
    $o
}

$cases = @(
    @{ Name = 'Old, no EPP match -> process';        Item = (New-Item2 'Prod\' 400);  Epp = @{ Status = 'Not Found'; Bypass = $true }; Expect = $true }
    @{ Name = 'Old, EPP Live -> keep';               Item = (New-Item2 'Prod\' 400);  Epp = @{ Status = 'Live'; Bypass = $true };      Expect = $false }
    @{ Name = 'Old, EPP lookup error -> keep';       Item = (New-Item2 'Prod\' 400);  Epp = @{ Status = 'Error'; Bypass = $true };     Expect = $false }
    @{ Name = 'Young, no EPP match -> keep';         Item = (New-Item2 'Prod\' 10);   Epp = @{ Status = 'Not Found'; Bypass = $true }; Expect = $false }
    @{ Name = 'Staging 100d, no EPP -> process';     Item = (New-Item2 'Staging\' 100); Epp = @{ Status = 'Not Found'; Bypass = $true }; Expect = $true }
    @{ Name = 'Staging 50d -> keep';                 Item = (New-Item2 'Staging\' 50); Epp = @{ Status = 'Not Found'; Bypass = $true }; Expect = $false }
    @{ Name = 'Retired EPP -> process';              Item = (New-Item2 'Prod\' 5);    Epp = @{ Status = 'Retired'; Bypass = $false };   Expect = $true }
    @{ Name = 'Retired EPP but limiting -> keep';    Item = (New-Item2 'Prod\' 5 @{ LimitingCollection = 'True' }); Epp = @{ Status = 'Retired'; Bypass = $false }; Expect = $false }
    @{ Name = 'Old but referenced -> keep';          Item = (New-Item2 'Prod\' 400 @{ ReferencedCollection = 'True' }); Epp = @{ Status = 'Not Found'; Bypass = $true }; Expect = $false }
    @{ Name = 'Old but app in TS -> keep';           Item = (New-Item2 'Prod\' 400 @{ 'App TS References' = 'True' }); Epp = @{ Status = 'Not Found'; Bypass = $true }; Expect = $false }
    @{ Name = 'Old but Adaptiva ref -> keep';        Item = (New-Item2 'Prod\' 400 @{ 'Adaptiva Policy Collection Reference' = $true }); Epp = @{ Status = 'Not Found'; Bypass = $true }; Expect = $false }
)

$fail = 0
foreach ($c in $cases) {
    $r = Test-CollectionProcessingRules -CollectionItem $c.Item -EPPStatus $c.Epp -ADGroupInfo $noAd -ReferenceDate $ref -Logger $logger
    $ok = $r.ShouldProcess -eq $c.Expect
    if (-not $ok) { $fail++ }
    '{0}  {1}  ({2})' -f $(if ($ok) { 'PASS' } else { 'FAIL' }), $c.Name, $r.Reason
}

# Empty EPP table must be reported as unverifiable, not "Not Found"
$empty = New-Object System.Data.DataTable
foreach ($col in 'EPP Status', 'Package Name', 'Package Short Name') { [void]$empty.Columns.Add($col, [string]) }
$s = Get-EPPPackageStatus -EPPData $empty -PackageName 'Some App' -AppComments '' -Logger $logger
$ok = $s.Status -eq 'Error'
if (-not $ok) { $fail++ }
'{0}  Empty EPP table -> Error ({1})' -f $(if ($ok) { 'PASS' } else { 'FAIL' }), $s.Status

function New-EppTable($rows) {
    $t = New-Object System.Data.DataTable
    foreach ($col in 'EPP Status', 'Package Name', 'Package Short Name') { [void]$t.Columns.Add($col, [string]) }
    foreach ($r in $rows) { [void]$t.Rows.Add($r[0], $r[1], $r[2]) }
    , $t
}

function Assert-Epp($name, $table, $pkg, $comments, $expectStatus, $expectType = $null) {
    $script:EPPLookupCache = $null
    $s = Get-EPPPackageStatus -EPPData $table -PackageName $pkg -AppComments $comments -Logger $logger
    $ok = ($s.Status -eq $expectStatus) -and ($null -eq $expectType -or $s.MatchType -eq $expectType)
    if (-not $ok) { $script:fail++ }
    '{0}  {1}  (got {2} / {3})' -f $(if ($ok) { 'PASS' } else { 'FAIL' }), $name, $s.Status, $s.MatchType
}

# Item 7: names containing wildcard characters must match literally
$t = New-EppTable @(, @('Retired', '[Gen] Fabrikam Client 2.3 #73', ''))
Assert-Epp 'Bracketed name exact' $t '[Gen] Fabrikam Client 2.3 #73' '' 'Retired' 'Package Name (Exact)'
Assert-Epp 'Bracketed substring (literal)' $t '[Gen] Fabrikam' '' 'Retired' 'Package Name (Contains)'
Assert-Epp 'Wildcard chars do not match everything' $t '[A-Z]' '' 'Not Found'
Assert-Epp 'Star is not a wildcard' $t '*' '' 'Not Found'

# Item 8: duplicates resolve to the most restrictive status regardless of row order
$t = New-EppTable @( @('Live', 'Backup Agent 5.5 #175', 'BKP_1'), @('Retired', 'Backup Agent 5.5 #175', 'BKP_1') )
Assert-Epp 'Duplicate name: Live row first' $t 'Backup Agent 5.5 #175' '' 'Live'
$t = New-EppTable @( @('Retired', 'Backup Agent 5.5 #175', 'BKP_1'), @('Live', 'Backup Agent 5.5 #175', 'BKP_1') )
Assert-Epp 'Duplicate name: Retired row first' $t 'Backup Agent 5.5 #175' '' 'Live'
Assert-Epp 'Duplicate short name' $t 'x' 'BKP_1' 'Live'

# Item 8: several "contains" hits resolve to the most restrictive
$t = New-EppTable @( @('Retired', 'Reporting Agent 1.0', ''), @('Live', 'Reporting Agent 1.0 Plus', '') )
Assert-Epp 'Contains: active beats retired' $t 'Reporting Agent' '' 'Live' 'Package Name (Contains)'

# Item 10: case and whitespace handled the same way for every caller
$t = New-EppTable @(, @('Cancel Package', '  Network Utility 7.3 #123 ', ''))
Assert-Epp 'Case-insensitive, trimmed exact match' $t 'network utility 7.3 #123' '' 'Cancel Package' 'Package Name (Exact)'

# Item 10: task sequence matcher agrees with the collection matcher
$script:EPPLookupCache = $null
$ts = New-Object System.Data.DataTable
foreach ($col in 'Task Sequence Name','Task Sequence Package ID','Task Sequence Object Path','Referenced Package Name','ReferencePackageID','Collection Name','CollectionID','Collection Type','Collection Object Path') { [void]$ts.Columns.Add($col, [string]) }
[void]$ts.Rows.Add('TS1','P1','\','network utility 7.3 #123','R1','','','','')
[void]$ts.Rows.Add('TS2','P2','\','Totally Unknown','R2','','','','')
$lk = Get-EPPLookup -EPPData $t -Logger $logger
$m = @(Get-TaskSequenceEPPMatches -TaskSequenceData $ts -Lookup $lk -Logger $logger)
$ok = ($m.Count -eq 1) -and ($m[0].'EPP Status' -eq 'Cancel Package')
if (-not $ok) { $fail++ }
'{0}  TS matcher uses shared logic ({1} match)' -f $(if ($ok) { 'PASS' } else { 'FAIL' }), $m.Count

# Item 9: lookup is built once per table
$script:EPPLookupCache = $null
$a = Get-EPPLookup -EPPData $t -Logger $logger
$b = Get-EPPLookup -EPPData $t -Logger $logger
$ok = [object]::ReferenceEquals($a, $b)
if (-not $ok) { $fail++ }
'{0}  Lookup cached per EPP table' -f $(if ($ok) { 'PASS' } else { 'FAIL' })

# Item 11: apostrophes are stored unchanged
$ok = (ConvertTo-SafeSqlString -Value "O'Brien Tools") -eq "O'Brien Tools"
if (-not $ok) { $fail++ }
'{0}  Apostrophe not doubled' -f $(if ($ok) { 'PASS' } else { 'FAIL' })
$ok = (ConvertTo-SafeSqlString -Value $null -DefaultValue 'N/A') -eq 'N/A'
if (-not $ok) { $fail++ }
'{0}  Null falls back to default' -f $(if ($ok) { 'PASS' } else { 'FAIL' })

# Stage 3: -days drives the age threshold
$script:NonStagingDays = 30
$nf = @{ Status = 'Not Found'; Bypass = $true }
foreach ($t in @(
    @{ Name = '-days 30: non-staging 40d -> process'; Item = (New-Item2 'Prod\' 40); Expect = $true },
    @{ Name = '-days 30: non-staging 20d -> keep';    Item = (New-Item2 'Prod\' 20); Expect = $false },
    @{ Name = '-days 30: staging 40d -> process';     Item = (New-Item2 'Staging\' 40); Expect = $true },
    @{ Name = '-days 30: staging 20d -> keep';        Item = (New-Item2 'Staging\' 20); Expect = $false }
)) {
    $r = Test-CollectionProcessingRules -CollectionItem $t.Item -EPPStatus $nf -ADGroupInfo $noAd -ReferenceDate $ref -Logger $logger
    $ok = $r.ShouldProcess -eq $t.Expect
    if (-not $ok) { $fail++ }
    '{0}  {1}' -f $(if ($ok) { 'PASS' } else { 'FAIL' }), $t.Name
}
$script:NonStagingDays = $null

# Stage 4: task sequence EPP cleanup selection
$m = @(
    [PSCustomObject]@{ 'Task Sequence Package ID' = 'TS1'; 'EPP Status' = 'Retired' },
    [PSCustomObject]@{ 'Task Sequence Package ID' = 'TS1'; 'EPP Status' = 'Live' },
    [PSCustomObject]@{ 'Task Sequence Package ID' = 'TS2'; 'EPP Status' = 'Retired' },
    [PSCustomObject]@{ 'Task Sequence Package ID' = 'TS2'; 'EPP Status' = 'Cancel Package' },
    [PSCustomObject]@{ 'Task Sequence Package ID' = 'TS3'; 'EPP Status' = 'Retired' },
    [PSCustomObject]@{ 'Task Sequence Package ID' = 'TS4'; 'EPP Status' = 'Not Found' }
)
$sel = @(Select-RetiredTaskSequenceMatches -Matches $m -ExcludePackageIds @('TS3'))
$ids = @($sel | ForEach-Object { $_.'Task Sequence Package ID' } | Sort-Object -Unique)
$ok = ($ids -join ',') -eq 'TS2' -and $sel.Count -eq 2
if (-not $ok) { $fail++ }
'{0}  TS cleanup: only all-retired, not in retired folder ({1})' -f $(if ($ok) { 'PASS' } else { 'FAIL' }), ($ids -join ',')

"Failures: $fail"
exit $fail
