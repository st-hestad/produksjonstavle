# =====================================================================
# VIKTIG: API-bruksvilkar
# =====================================================================
# Dette scriptet gjor sporringar mot Landax sitt REST-API.
# Det bor IKKE kjores oftere enn avtalt med Landax.
# Anbefalt testintervall: minst 30 minutter mellom hvert kjoret.
# Hyppigere eller live oppdatering ma avklaras med Landax
# for scriptet tas i produksjonsbruk med korte intervaller.
# =====================================================================

param(
[switch]$FullScan,
[switch]$OptimizedApi
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$requiredVars = @("LANDAX_BASEURL", "LANDAX_USERNAME", "LANDAX_PASSWORD")
$missingVars = $requiredVars | Where-Object { [string]::IsNullOrWhiteSpace((Get-Item "env:$_" -ErrorAction SilentlyContinue).Value) }

if ($missingVars) {
    $envFile = Join-Path $PSScriptRoot "../Miljøvariabler/set-env-prod.ps1"
    if (Test-Path $envFile) {
        . $envFile
    }
    $missingVars = $requiredVars | Where-Object { [string]::IsNullOrWhiteSpace((Get-Item "env:$_" -ErrorAction SilentlyContinue).Value) }
    if ($missingVars) {
        throw "Følgende miljøvariabler mangler: $($missingVars -join ', ')"
    }
}

$configPath = "$PSScriptRoot/../02_config/landax-production-board.config.ps1"
if (-not (Test-Path $configPath)) {
throw "Mangler config-fil: $configPath"
}
. "$PSScriptRoot/../02_config/landax-production-board.config.ps1"

$SharePointPublishPath = "/Users/sverretobiashestad/Library/CloudStorage/OneDrive-MekonAS/Team Mekon - Infoskjerm"
$HtmlOutputPath = $script:HtmlOutputPath

$scriptStart = Get-Date

function Write-StepTiming([string]$Name, [datetime]$Start, [datetime]$End) {
$duration = [Math]::Round(($End - $Start).TotalSeconds, 1)
Write-Host "$Name - Start: $($Start.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host "$Name - Slutt: $($End.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host ("{0}: {1}s" -f $Name, $duration)
}

$LandaxBaseUrl = $env:LANDAX_BASEURL
$LandaxUsername = $env:LANDAX_USERNAME
$LandaxPassword = $env:LANDAX_PASSWORD

$landaxBaseUrl = if ($LandaxBaseUrl) { $LandaxBaseUrl.TrimEnd('/') } else { $script:LandaxBaseUrlDefault }
$landaxUsername = $LandaxUsername
$landaxPassword = $LandaxPassword
if (-not $landaxUsername -or -not $landaxPassword) {
throw "LANDAX_USERNAME og/eller LANDAX_PASSWORD mangler."
}

# --- Auth & HTTP-hjelpere ---

$script:apiCallCount = 0

function New-BasicAuthHeader([string]$username, [string]$password) {
$basic = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${username}`:${password}"))
return @{ Authorization = "Basic $basic"; Accept = "application/json" }
}

function Invoke-LandaxGet([string]$url, [hashtable]$headers) {
$script:apiCallCount++
$resp = Invoke-WebRequest -Method GET -Uri $url -Headers $headers -SkipHttpErrorCheck -HttpVersion 1.1
$statusCode = [int]$resp.StatusCode
if ($statusCode -lt 200 -or $statusCode -ge 300) {
throw "HTTP $statusCode ved GET $url"
}
if ([string]::IsNullOrWhiteSpace($resp.Content)) { return $null }
return ($resp.Content | ConvertFrom-Json -Depth 100)
}

function Get-LandaxAll([string]$endpoint, [hashtable]$headers) {
$all = @()
$from = 0
$count = 500
$seenIds = @{}
$page = 0
$maxPages = 200

while ($true) {
$page++
if ($page -gt $maxPages) { break }

$url = if ($endpoint -match '\?') {
"$landaxBaseUrl/${endpoint}&from=$from&count=$count"
} else {
"$landaxBaseUrl/${endpoint}?from=$from&count=$count"
}
$resp = Invoke-LandaxGet -url $url -headers $headers
if ($null -eq $resp) { break }

$batch = if ($resp.PSObject.Properties.Name -contains "value") { @($resp.value) } else { @($resp) }
if ($batch.Count -eq 0) { break }

$newInBatch = 0
foreach ($item in $batch) {
if ($null -ne $item -and $item.PSObject.Properties.Name -contains "Id" -and $null -ne $item.Id) {
$key = [string]$item.Id
if (-not $seenIds.ContainsKey($key)) {
$seenIds[$key] = $true
$all += $item
$newInBatch++
}
} else {
$all += $item
$newInBatch++
}
}

if ($newInBatch -eq 0 -or $batch.Count -lt $count) { break }
$from += $count
}

return $all
}

# Sjekk om $select stottes pa et endepunkt
# Returnerer $true hvis svaret respekterer $select (antall felter redusert)
# og $false hvis alle felt returneres uansett.
function Test-SelectSupport([string]$Endpoint, [hashtable]$Headers, [string]$SelectFields) {
try {
$url = "$landaxBaseUrl/${Endpoint}?from=0&count=1&`$select=$SelectFields"
$resp = Invoke-LandaxGet -url $url -headers $Headers
if ($null -eq $resp) { return $false }
$item = if ($resp.PSObject.Properties.Name -contains "value") { @($resp.value)[0] } else { $resp }
if ($null -eq $item) { return $false }
$expectedFields = @($SelectFields -split "," | ForEach-Object { $_.Trim() })
$returnedFields = @($item.PSObject.Properties.Name)
# Stottes hvis antall returnerte felt er lik eller faerre enn forespurt + noen overhead-felt (odata.type osv.)
return ($returnedFields.Count -le ($expectedFields.Count + 3))
} catch { return $false }
}

# Bygg URL-parameter for $select (URL-enkoder komma, henter riktig)
function Build-SelectParam([string]$Fields) {
return "`$select=$Fields"
}

$landaxHeaders = New-BasicAuthHeader -username $landaxUsername -password $landaxPassword

# --- Snapshot setup ---

$logDir = $script:ProductionBoardLogDir
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
$snapshotPath = $script:SnapshotPath
$perfPath = $script:PerfPath
$baselinePerfPath = $script:BaselinePerfPath

$previousSnapshot = $null
if (Test-Path $snapshotPath) {
try { $previousSnapshot = Get-Content -Raw -Path $snapshotPath | ConvertFrom-Json -Depth 100 } catch {}
}

$baselineSeconds = $null
$baselineApiCalls = $null
if (Test-Path $baselinePerfPath) {
try {
$baselinePerf = Get-Content -Raw -Path $baselinePerfPath | ConvertFrom-Json -Depth 10
if ($baselinePerf.PSObject.Properties.Name -contains "seconds") { $baselineSeconds = [double]$baselinePerf.seconds }
if ($baselinePerf.PSObject.Properties.Name -contains "apiCalls") { $baselineApiCalls = [int]$baselinePerf.apiCalls }
} catch {}
}

# --- Tidssone og uke-grenser ---

$osloTZ       = [System.TimeZoneInfo]::FindSystemTimeZoneById("Europe/Oslo")
$nowOslo      = [System.TimeZoneInfo]::ConvertTime((Get-Date), $osloTZ)
$todayOslo    = $nowOslo.Date
$recentCutoff = $nowOslo.AddHours(-8)
$mondayOffset = (([int]$todayOslo.DayOfWeek + 6) % 7)
$weekStart    = $todayOslo.AddDays(-$mondayOffset)
$weekEnd      = $weekStart.AddDays(7)

function ConvertTo-OsloDateTime([object]$DateValue, [bool]$AssumeUtc = $true) {
if ($null -eq $DateValue -or [string]::IsNullOrWhiteSpace([string]$DateValue)) { return $null }
$raw = [string]$DateValue
if ($raw -match '^\d{4}-\d{2}-\d{2}$') {
return [datetime]::ParseExact($raw, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
}
$dt = [datetime]$DateValue
if ($dt.Kind -eq [System.DateTimeKind]::Utc) {
return [System.TimeZoneInfo]::ConvertTimeFromUtc($dt, $osloTZ)
}
if ($dt.Kind -eq [System.DateTimeKind]::Local) {
return [System.TimeZoneInfo]::ConvertTime($dt, $osloTZ)
}
if ($AssumeUtc) {
return [System.TimeZoneInfo]::ConvertTimeFromUtc([datetime]::SpecifyKind($dt, [System.DateTimeKind]::Utc), $osloTZ)
}
return $dt
}

function ConvertTo-OsloDate([object]$DateValue, [bool]$AssumeUtc = $true) {
$dt = ConvertTo-OsloDateTime -DateValue $DateValue -AssumeUtc $AssumeUtc
if ($null -eq $dt) { return $null }
return $dt.Date
}

function Get-NorwegianPublicHolidays([int]$Year) {
$a = $Year % 19
$b = [int][Math]::Floor($Year / 100)
$c = $Year % 100
$d = [int][Math]::Floor($b / 4)
$e = $b % 4
$f = [int][Math]::Floor(($b + 8) / 25)
$g = [int][Math]::Floor(($b - $f + 1) / 3)
$h = (19 * $a + $b - $d - $g + 15) % 30
$i = [int][Math]::Floor($c / 4)
$k = $c % 4
$l = (32 + 2 * $e + 2 * $i - $h - $k) % 7
$m = [int][Math]::Floor(($a + 11 * $h + 22 * $l) / 451)
$easterMonth = [int][Math]::Floor(($h + $l - 7 * $m + 114) / 31)
$easterDay = (($h + $l - 7 * $m + 114) % 31) + 1
$easterSunday = (Get-Date -Year $Year -Month $easterMonth -Day $easterDay).Date

$holidays = New-Object System.Collections.Generic.List[object]

$holidays.Add([PSCustomObject]@{ Date = (Get-Date -Year $Year -Month 1 -Day 1).Date; Name = "Nyttårsdag" }) | Out-Null
$holidays.Add([PSCustomObject]@{ Date = (Get-Date -Year $Year -Month 5 -Day 1).Date; Name = "Arbeidernes dag" }) | Out-Null
$holidays.Add([PSCustomObject]@{ Date = (Get-Date -Year $Year -Month 5 -Day 17).Date; Name = "Grunnlovsdag" }) | Out-Null
$holidays.Add([PSCustomObject]@{ Date = (Get-Date -Year $Year -Month 12 -Day 25).Date; Name = "1. juledag" }) | Out-Null
$holidays.Add([PSCustomObject]@{ Date = (Get-Date -Year $Year -Month 12 -Day 26).Date; Name = "2. juledag" }) | Out-Null

$holidays.Add([PSCustomObject]@{ Date = $easterSunday.AddDays(-3).Date; Name = "Skjærtorsdag" }) | Out-Null
$holidays.Add([PSCustomObject]@{ Date = $easterSunday.AddDays(-2).Date; Name = "Langfredag" }) | Out-Null
$holidays.Add([PSCustomObject]@{ Date = $easterSunday.Date; Name = "1. påskedag" }) | Out-Null
$holidays.Add([PSCustomObject]@{ Date = $easterSunday.AddDays(1).Date; Name = "2. påskedag" }) | Out-Null
$holidays.Add([PSCustomObject]@{ Date = $easterSunday.AddDays(39).Date; Name = "Kristi himmelfartsdag" }) | Out-Null
$holidays.Add([PSCustomObject]@{ Date = $easterSunday.AddDays(49).Date; Name = "1. pinsedag" }) | Out-Null
$holidays.Add([PSCustomObject]@{ Date = $easterSunday.AddDays(50).Date; Name = "2. pinsedag" }) | Out-Null

return @($holidays | Sort-Object Date)
}

$holidayYears = @(
($nowOslo.Year - 1),
($nowOslo.Year),
($nowOslo.Year + 1),
($nowOslo.Year + 2)
) | Sort-Object -Unique
$norwegianPublicHolidays = @(
foreach ($holidayYear in $holidayYears) {
Get-NorwegianPublicHolidays -Year $holidayYear
}
)

$holidayNameByDate = @{}
foreach ($holiday in $norwegianPublicHolidays) {
$holidayDateKey = $holiday.Date.ToString("yyyy-MM-dd")
if (-not $holidayNameByDate.ContainsKey($holidayDateKey)) {
$holidayNameByDate[$holidayDateKey] = [string]$holiday.Name
}
}

$holidaysThisWeek = New-Object System.Collections.Generic.List[object]
$holidayCursor = $weekStart.Date
while ($holidayCursor -lt $weekEnd.Date) {
$weekDateKey = $holidayCursor.ToString("yyyy-MM-dd")
if ($holidayNameByDate.ContainsKey($weekDateKey)) {
$holidaysThisWeek.Add([PSCustomObject]@{
Date = $holidayCursor
Name = [string]$holidayNameByDate[$weekDateKey]
}) | Out-Null
}
$holidayCursor = $holidayCursor.AddDays(1)
}

Write-Host ("Antall røddager generert: {0}" -f $holidayNameByDate.Count)
Write-Host ("Røddager denne uke: {0}" -f $holidaysThisWeek.Count)
foreach ($holidayWeekItem in @($holidaysThisWeek | Sort-Object Date)) {
Write-Host ("- {0}: {1}" -f $holidayWeekItem.Date.ToString("yyyy-MM-dd"), [string]$holidayWeekItem.Name)
}

function Get-CachedLandaxDataset {
param(
[string]$Endpoint,
[string]$CacheFilePath,
[hashtable]$Headers,
[int]$MaxAgeHours = 24,
[string]$SelectFields = ""
)

$useCache = $false
if (Test-Path $CacheFilePath) {
$cacheAgeHours = ((Get-Date) - (Get-Item $CacheFilePath).LastWriteTime).TotalHours
if ($cacheAgeHours -lt $MaxAgeHours) {
$useCache = $true
}
}

if ($useCache) {
try {
$cached = Get-Content -Raw -Path $CacheFilePath | ConvertFrom-Json -Depth 100
if ($null -ne $cached) {
if ($cached -is [System.Array]) {
return [PSCustomObject]@{ Data = @($cached); Source = "Cache" }
}
return [PSCustomObject]@{ Data = @($cached); Source = "Cache" }
}
} catch {
# Ved lese/parsing-feil faller vi tilbake til API
}
}

$endpointWithSelect = $Endpoint
if (-not [string]::IsNullOrWhiteSpace($SelectFields)) {
$endpointWithSelect = "${Endpoint}?`$select=${SelectFields}"
}
$fresh = @(Get-LandaxAll -endpoint $endpointWithSelect -headers $Headers)
try {
$fresh | ConvertTo-Json -Depth 100 | Out-File -Path $CacheFilePath -Encoding utf8
} catch {
# Ignorer skrivefeil og fortsett med API-data i minnet
}

return [PSCustomObject]@{ Data = $fresh; Source = "API" }
}

function Test-TaskBaseFilter([object]$Task, [int]$ProduksjonDeptId) {
if ($null -eq $Task) { return $false }
if ([int]$Task.TypeId -ne 11) { return $false }
if ($null -eq $Task.DepartmentId -or [int]$Task.DepartmentId -ne $ProduksjonDeptId) { return $false }
if ($Task.IsTodo -ne $true) { return $false }
return $true
}

function Get-TaskAnchorDate([object]$Task) {
$anchors = @(
    (ConvertTo-OsloDateTime -DateValue $Task.ChangedDateTime -AssumeUtc $true)
    (ConvertTo-OsloDateTime -DateValue $Task.RegisteredDateTime -AssumeUtc $true)
    (ConvertTo-OsloDateTime -DateValue $Task.PlannedDoneDate -AssumeUtc $true)
    (ConvertTo-OsloDateTime -DateValue $Task.DoneDate -AssumeUtc $false)
) | Where-Object { $null -ne $_ }

if ($anchors.Count -eq 0) { return $null }
return (($anchors | Sort-Object -Descending)[0])
}

function Get-OptimizedTasks {
param(
[hashtable]$Headers,
[int]$ProduksjonDeptId,
[switch]$UseFullScan,
[datetime]$WeekStartDate,
[datetime]$RecentCutoffDate
)

if ($UseFullScan) {
return @(Get-LandaxAll -endpoint "tasks" -headers $Headers)
}

$pageSize = 200
$maxPages = 3
$relevanceThreshold = $WeekStartDate.AddDays(-14)

$all = [System.Collections.Generic.List[object]]::new()
$seenIds = @{}

function Add-UniqueTasks([array]$Batch, [switch]$ApplyLocalBaseFilter) {
foreach ($item in $Batch) {
if ($null -eq $item -or $null -eq $item.Id) { continue }
$k = [string]$item.Id
if ($seenIds.ContainsKey($k)) { continue }
if ($ApplyLocalBaseFilter -and -not (Test-TaskBaseFilter -Task $item -ProduksjonDeptId $ProduksjonDeptId)) { continue }
$seenIds[$k] = $true
$all.Add($item) | Out-Null
}
}

# Forsok server-side basefilter pa tasks-endepunkt
$from = 0
$serverFilterUrl = "$landaxBaseUrl/tasks?from=0&count=$pageSize&typeId=11&departmentId=$ProduksjonDeptId&isTodo=true"
$serverFilterResp = Invoke-LandaxGet -url $serverFilterUrl -headers $Headers
$serverFilterBatch = if ($null -ne $serverFilterResp) {
if ($serverFilterResp.PSObject.Properties.Name -contains "value") { @($serverFilterResp.value) } else { @($serverFilterResp) }
} else { @() }

$serverSideSupported = ($serverFilterBatch.Count -eq 0 -or @($serverFilterBatch | Where-Object { -not (Test-TaskBaseFilter -Task $_ -ProduksjonDeptId $ProduksjonDeptId) }).Count -eq 0)

if ($serverSideSupported) {
Add-UniqueTasks -Batch $serverFilterBatch
$from += $pageSize

for ($page = 2; $page -le $maxPages; $page++) {
$url = "$landaxBaseUrl/tasks?from=$from&count=$pageSize&typeId=11&departmentId=$ProduksjonDeptId&isTodo=true"
$resp = Invoke-LandaxGet -url $url -headers $Headers
$batch = if ($null -ne $resp) {
if ($resp.PSObject.Properties.Name -contains "value") { @($resp.value) } else { @($resp) }
} else { @() }

if ($batch.Count -eq 0) { break }
Add-UniqueTasks -Batch $batch

$allOld = $true
foreach ($t in $batch) {
$a = Get-TaskAnchorDate -Task $t
if ($null -eq $a -or $a -ge $relevanceThreshold) { $allOld = $false; break }
}
if ($allOld) { break }

if ($batch.Count -lt $pageSize) { break }
$from += $pageSize
}

return @($all)
}

# Fallback: hent nyeste tasks med paging-stopp, filtrer basekrav lokalt
$from = 0
for ($page = 1; $page -le $maxPages; $page++) {
$url = "$landaxBaseUrl/tasks?from=$from&count=$pageSize"
$resp = Invoke-LandaxGet -url $url -headers $Headers
$batch = if ($null -ne $resp) {
if ($resp.PSObject.Properties.Name -contains "value") { @($resp.value) } else { @($resp) }
} else { @() }

if ($batch.Count -eq 0) { break }
Add-UniqueTasks -Batch $batch -ApplyLocalBaseFilter

$allOld = $true
foreach ($t in $batch) {
$a = Get-TaskAnchorDate -Task $t
if ($null -eq $a -or $a -ge $relevanceThreshold) { $allOld = $false; break }
}
if ($allOld) { break }

if ($batch.Count -lt $pageSize) { break }
$from += $pageSize
}

return @($all)
}

# =====================================================================
# TRINN 1 - Hent nodvendige data
# =====================================================================

# $select-felter per endepunkt (minimalt nødvendige felt)
$selectDepartments  = "Id,Name"
$selectCoworkers    = "Id,DisplayName,FullName,IsActive"
$selectProjects     = "Id,Number,Name,ParentProjectId,IsMainProject"
$selectTasks        = "Id,Description,TypeId,DepartmentId,IsTodo,IsClosed,StatusId,PlannedStartDate,PlannedDoneDate,DoneDate,ChangedDateTime,RegisteredDateTime,HandledByCoworkerId,ProjectId,ClosestProjectId,MainProjectId,Progress"
$selectParticipants = "Id,TaskId,CoworkerId,Name,ObjectId,ModuleId,RecordId,SourceId,ParentId"

$departmentsCachePath = $script:DepartmentsCachePath
$projectsCachePath    = $script:ProjectsCachePath
$coworkersCachePath   = $script:CoworkersCachePath

# Test $select-stotte pa tasks (kun ved fersk henting, ikke fra cache)
$tasksSelectSupported = $false
$tasksSelectTested    = $false

$departmentsApiStart = Get-Date
$departmentsResult = Get-CachedLandaxDataset -Endpoint "departments" -CacheFilePath $departmentsCachePath -Headers $landaxHeaders -MaxAgeHours 24 -SelectFields $selectDepartments
$departments = @($departmentsResult.Data)
$departmentsApiEnd = Get-Date
Write-StepTiming -Name "Departments API" -Start $departmentsApiStart -End $departmentsApiEnd
Write-Host "Departments: $($departmentsResult.Source)"

$produksjonDeptId = $null
foreach ($d in $departments) {
if ([string]$d.Name -eq "Produksjon") { $produksjonDeptId = [int]$d.Id; break }
}
if ($null -eq $produksjonDeptId) { throw "Fant ikke avdeling 'Produksjon' i Landax." }

# Test om $select er støttet på tasks-endepunktet
$tasksSelectTested = $true
$tasksSelectSupported = Test-SelectSupport -Endpoint "tasks" -Headers $landaxHeaders -SelectFields $selectTasks
if ($tasksSelectSupported) {
Write-Host "tasks: `$select stottes — begrenser felter"
} else {
Write-Host "tasks: `$select stottes ikke pa dette endepunktet — henter alle felt"
}

function Get-TasksFastDefault([hashtable]$Headers) {
$selectParam = if ($tasksSelectSupported) { "&`$select=$selectTasks" } else { "" }
$url = "$landaxBaseUrl/tasks?from=0&count=1000${selectParam}"
$resp = Invoke-LandaxGet -url $url -headers $Headers
if ($null -eq $resp) { return @() }
if ($resp.PSObject.Properties.Name -contains "value") { return @($resp.value) }
return @($resp)
}

$effectiveFullScan = $FullScan.IsPresent
$effectiveOptimizedApi = $OptimizedApi.IsPresent
if (-not $effectiveFullScan -and -not $effectiveOptimizedApi) {
$effectiveFullScan = [bool]$script:DefaultFullScan
$effectiveOptimizedApi = [bool]$script:DefaultOptimizedApi
}

$tasksApiStart = Get-Date
$tasksApiMode = $script:DefaultApiMode
if ($effectiveOptimizedApi) {
$tasksApiMode = "OptimizedApi + LocalFilter + Select"
$tasks = @(Get-OptimizedTasks -Headers $landaxHeaders -ProduksjonDeptId $produksjonDeptId -UseFullScan:$effectiveFullScan -WeekStartDate $weekStart -RecentCutoffDate $recentCutoff)
} elseif ($effectiveFullScan) {
$tasksApiMode = "FullScan + Select"
$selectEp = if ($tasksSelectSupported) { "tasks?`$select=$selectTasks" } else { "tasks" }
$tasks = @(Get-LandaxAll -endpoint $selectEp -headers $landaxHeaders)
} else {
$tasksApiMode = "FullScan + LocalFilter + Select"
$tasks = @(Get-TasksFastDefault -Headers $landaxHeaders)
}
$tasksApiEnd = Get-Date
Write-StepTiming -Name "Tasks API" -Start $tasksApiStart -End $tasksApiEnd

$coworkersApiStart = Get-Date
$coworkersResult = Get-CachedLandaxDataset -Endpoint "coworkers" -CacheFilePath $coworkersCachePath -Headers $landaxHeaders -MaxAgeHours 24 -SelectFields $selectCoworkers
$coworkers   = @($coworkersResult.Data)
$coworkersApiEnd = Get-Date
Write-StepTiming -Name "Coworkers API" -Start $coworkersApiStart -End $coworkersApiEnd
Write-Host "Coworkers: $($coworkersResult.Source)"


$projectsApiStart = Get-Date
$projectsResult = Get-CachedLandaxDataset -Endpoint "projects" -CacheFilePath $projectsCachePath -Headers $landaxHeaders -MaxAgeHours 24 -SelectFields $selectProjects
$projects    = @($projectsResult.Data)
$projectsApiEnd = Get-Date
Write-StepTiming -Name "Projects API" -Start $projectsApiStart -End $projectsApiEnd
Write-Host "Projects: $($projectsResult.Source)"

$baseApiCalls = $script:apiCallCount
Write-Host "Tasks hentet: $($tasks.Count)"
Write-Host "Tasks før filter: $($tasks.Count)"
Write-Host "Tasks API-modus: $tasksApiMode"

$taskStatusNameById = @{}
$taskStatusLookupEndpoint = $null
foreach ($statusEndpoint in @("taskstatuses", "taskstatus", "statuses")) {
try {
$statusItems = @(Get-LandaxAll -endpoint $statusEndpoint -headers $landaxHeaders)
if ($statusItems.Count -eq 0) { continue }
foreach ($statusItem in $statusItems) {
if ($null -eq $statusItem) { continue }

$statusId = $null
foreach ($idField in @("Id", "StatusId")) {
if ($statusItem.PSObject.Properties.Name -contains $idField -and $null -ne $statusItem.$idField) {
try { $statusId = [int]$statusItem.$idField; break } catch {}
}
}
if ($null -eq $statusId) { continue }

$statusName = $null
foreach ($nameField in @("Name", "DisplayName", "Description", "Title")) {
if ($statusItem.PSObject.Properties.Name -contains $nameField -and -not [string]::IsNullOrWhiteSpace([string]$statusItem.$nameField)) {
$statusName = [string]$statusItem.$nameField
break
}
}

if (-not [string]::IsNullOrWhiteSpace($statusName) -and -not $taskStatusNameById.ContainsKey($statusId)) {
$taskStatusNameById[$statusId] = $statusName.Trim()
}
}

if ($taskStatusNameById.Count -gt 0) {
$taskStatusLookupEndpoint = $statusEndpoint
break
}
} catch {
# Status-endepunkt kan mangle i noen Landax-oppsett; fallback brukes senere.
}
}

if ($taskStatusNameById.Count -gt 0) {
Write-Host ("Task-status lookup: fant {0} statusnavn fra endpoint '{1}'" -f $taskStatusNameById.Count, $taskStatusLookupEndpoint)
} else {
Write-Host "Task-status lookup: ingen statusnavn-endepunkt tilgjengelig; bruker task-felt/fallback"
}

$taskStatusNameByIdLocal = @{
1 = "Ikke påbegynt"
2 = "Påbegynt"
}

foreach ($statusIdKey in $taskStatusNameByIdLocal.Keys) {
if (-not $taskStatusNameById.ContainsKey([int]$statusIdKey)) {
$taskStatusNameById[[int]$statusIdKey] = [string]$taskStatusNameByIdLocal[$statusIdKey]
}
}

Write-Host ("Task-status lookup: totale statusnavn etter lokal fallback: {0}" -f $taskStatusNameById.Count)

# =====================================================================
# TRINN 4 - Bygg raske oppslagstabeller
# =====================================================================

$coworkerById = @{}
foreach ($c in $coworkers) {
if ($null -ne $c.Id) { $coworkerById[[int]$c.Id] = $c.DisplayName }
}

$projectById       = @{}
$projectNumberById = @{}
$projectParentById = @{}
$projectIsMainById = @{}
foreach ($p in $projects) {
if ($null -ne $p.Id) {
$projId = [int]$p.Id
$projectById[$projId]       = $p.Name
$projectNumberById[$projId] = $p.Number
if ($null -ne $p.ParentProjectId -and $p.ParentProjectId -ne '') {
$projectParentById[$projId] = [int]$p.ParentProjectId
}
$projectIsMainById[$projId] = ($p.IsMainProject -eq $true)
}
}

function Get-DisplayProjectId([object]$ClosestId, [object]$ProjectId) {
$startId = $null
if ($null -ne $ClosestId -and $ClosestId -ne '') { $startId = [int]$ClosestId }
elseif ($null -ne $ProjectId -and $ProjectId -ne '') { $startId = [int]$ProjectId }
if ($null -eq $startId -or -not $projectById.ContainsKey($startId)) { return $null }

$current = $startId
while ($projectParentById.ContainsKey($current)) {
$parent = $projectParentById[$current]
if ($projectIsMainById.ContainsKey($parent) -and $projectIsMainById[$parent]) { break }
$current = $parent
}
return $current
}

# =====================================================================
# TRINN 2 - Filtrer tasks umiddelbart
# =====================================================================

$taskFilteringStart = Get-Date
$landaxStatusNameCount = 0
$fallbackStatusCount = 0
$statusIdWithoutNameCount = 0
$statusIdWithoutNameSamples = New-Object System.Collections.Generic.List[string]
$filteredTasks = @(foreach ($t in $tasks) {
if ([int]$t.TypeId -ne 11) { continue }
if ($null -eq $t.DepartmentId -or [int]$t.DepartmentId -ne $produksjonDeptId) { continue }
if ($t.IsTodo -ne $true) { continue }

$isClosed = ($t.IsClosed -eq $true)

# Beregn board-datoer
$boardStart = $null
$boardEnd   = $null
$plannedStartLocal = ConvertTo-OsloDate -DateValue $t.PlannedStartDate -AssumeUtc $true
$plannedDoneLocal  = ConvertTo-OsloDate -DateValue $t.PlannedDoneDate  -AssumeUtc $true
$doneDateLocal     = ConvertTo-OsloDateTime -DateValue $t.DoneDate     -AssumeUtc $false

if ($isClosed) {
$boardStart = $plannedStartLocal
$boardEnd   = $plannedDoneLocal
if ($null -eq $boardStart) { $boardStart = $boardEnd }
if ($null -eq $boardEnd)   { $boardEnd   = $boardStart }
if ($null -eq $boardStart -and $null -ne $doneDateLocal) {
$boardStart = $doneDateLocal.Date
$boardEnd   = $doneDateLocal.Date
}
} else {
$boardStart = $plannedStartLocal
$boardEnd   = $plannedDoneLocal
if ($null -eq $boardStart) { $boardStart = $boardEnd }
if ($null -eq $boardEnd)   { $boardEnd   = $boardStart }
}

if ($null -ne $boardStart -and $null -ne $boardEnd -and $boardEnd -lt $boardStart) {
$tmp = $boardStart
$boardStart = $boardEnd
$boardEnd = $tmp
}

# Ukefilter basert på intervall-overlapp med eksklusiv WeekEnd
$overlapsWeek = ($null -ne $boardStart -and $null -ne $boardEnd -and $boardStart -lt $weekEnd -and $boardEnd -ge $weekStart)

if ([int]$t.Id -eq 3420) {
Write-Host "Task 3420"
Write-Host ("Start: {0}" -f $boardStart)
Write-Host ("Slutt: {0}" -f $boardEnd)
Write-Host ("WeekStart: {0}" -f $weekStart)
Write-Host ("WeekEnd: {0}" -f $weekEnd)
Write-Host ("PassUke: {0}" -f $overlapsWeek)
}

if ($isClosed) {
$doneLocal    = ConvertTo-OsloDateTime -DateValue $t.DoneDate        -AssumeUtc $false
$changedLocal = ConvertTo-OsloDateTime -DateValue $t.ChangedDateTime -AssumeUtc $true
$isDoneToday  = ($null -ne $doneLocal    -and $doneLocal.Date -eq $todayOslo)
$isRecent     = ($null -ne $changedLocal -and $changedLocal -ge $recentCutoff)
if (-not ($isDoneToday -or $isRecent)) { continue }
} else {
if (-not $overlapsWeek) { continue }
}

# Status
$plannedEndOslo = ConvertTo-OsloDate -DateValue $t.PlannedDoneDate -AssumeUtc $true
$landaxStatusName = $null
$statusIdValue = $null
if ($t.PSObject.Properties.Name -contains "StatusName" -and -not [string]::IsNullOrWhiteSpace([string]$t.StatusName)) {
$landaxStatusName = [string]$t.StatusName
} elseif ($t.PSObject.Properties.Name -contains "Status" -and -not [string]::IsNullOrWhiteSpace([string]$t.Status)) {
$landaxStatusName = [string]$t.Status
} elseif ($t.PSObject.Properties.Name -contains "StatusId" -and $null -ne $t.StatusId) {
try {
$statusIdValue = [int]$t.StatusId
if ($taskStatusNameById.ContainsKey($statusIdValue)) {
$landaxStatusName = [string]$taskStatusNameById[$statusIdValue]
}
} catch {}
}

if ($null -ne $statusIdValue -and [string]::IsNullOrWhiteSpace($landaxStatusName)) {
$statusIdWithoutNameCount++
if ($statusIdWithoutNameSamples.Count -lt 20) {
$statusIdWithoutNameSamples.Add(("TaskId={0},StatusId={1}" -f [int]$t.Id, $statusIdValue)) | Out-Null
}
}

$fallbackStatus = if ($isClosed) { "Ferdig" }
                  elseif ($null -ne $t.Progress -and [double]$t.Progress -gt 0 -and [double]$t.Progress -lt 100) { "Påbegynt" }
                  else { "Planlagt" }

$landaxStatusNormalized = if (-not [string]::IsNullOrWhiteSpace($landaxStatusName)) { $landaxStatusName.Trim().ToLowerInvariant() } else { "" }
$isClosedByLandaxStatus = $landaxStatusNormalized -in @("ferdig", "lukket", "closed", "completed")

$baseStatus = if ($isClosed -or $isClosedByLandaxStatus) {
"Ferdig"
} elseif ($landaxStatusNormalized -eq "ikke påbegynt") {
"Planlagt"
} elseif ($landaxStatusNormalized -eq "påbegynt" -or $landaxStatusNormalized -eq "pabegynt") {
"Påbegynt"
} elseif (-not [string]::IsNullOrWhiteSpace($landaxStatusName)) {
$landaxStatusName.Trim()
} else {
$fallbackStatus
}

if (-not [string]::IsNullOrWhiteSpace($landaxStatusName)) {
$landaxStatusNameCount++
} else {
$fallbackStatusCount++
}

$status = $baseStatus

# Prosjekt (hierarki-oppslag)
$dpId         = Get-DisplayProjectId -ClosestId $t.ClosestProjectId -ProjectId $t.ProjectId
$prosjektNavn = if ($null -ne $dpId -and $projectById.ContainsKey($dpId))       { $projectById[$dpId] }       else { $null }
$prosjektNr   = if ($null -ne $dpId -and $projectNumberById.ContainsKey($dpId)) { $projectNumberById[$dpId] } else { $null }

[PSCustomObject]@{
Id                  = $t.Id
HandledByCoworkerId = $t.HandledByCoworkerId
Oppgave             = $t.Description
BoardStart          = $boardStart
BoardEnd            = $boardEnd
Prosjekt            = $prosjektNavn
ProsjektNummer      = $prosjektNr
Lukket              = $isClosed
Status              = $status
LandaxStatusRaw     = $landaxStatusName
PlannedStartLocal   = $plannedStartLocal
PlannedDoneLocal    = $plannedDoneLocal
DoneDateLocal       = $doneDateLocal
}
})
$taskFilteringEnd = Get-Date
Write-StepTiming -Name "Task-filtrering" -Start $taskFilteringStart -End $taskFilteringEnd

Write-Host "Tasks etter filter: $($filteredTasks.Count)"
Write-Host ("Task-IDer etter filter: " + ((@($filteredTasks | ForEach-Object { [string]$_.Id } | Sort-Object -Unique) -join ", ")))
Write-Host ("Oppgaver med Landax-statusnavn: {0}" -f $landaxStatusNameCount)
Write-Host ("Oppgaver med fallback-status: {0}" -f $fallbackStatusCount)
Write-Host ("Oppgaver med StatusId uten statusnavn-oppslag (terminaldiagnostikk): {0}" -f $statusIdWithoutNameCount)
if ($statusIdWithoutNameSamples.Count -gt 0) {
Write-Host ("StatusId-diagnostikk (utvalg): " + ($statusIdWithoutNameSamples -join "; "))
}

# Logg antall felter per task
if ($tasks.Count -gt 0) {
$sampleTaskFieldCount = @($tasks[0].PSObject.Properties).Count
$selectInfo = if ($tasksSelectSupported) { 'med $select' } else { 'uten $select' }
Write-Host "Felter per task: $sampleTaskFieldCount ($selectInfo)"
}

# =====================================================================
# TRINN 5 - Participants (bulk-henting, lokal filtrering)
# =====================================================================

$participantsByTaskId = @{}
$participantsApiStart = Get-Date

# Test $select-stotte pa participants
$participantsSelectSupported = Test-SelectSupport -Endpoint "participants" -Headers $landaxHeaders -SelectFields $selectParticipants
if ($participantsSelectSupported) {
Write-Host "participants: `$select stottes — begrenser felter"
} else {
Write-Host "participants: `$select stottes ikke pa dette endepunktet — henter alle felt"
}

$participantsEndpoint = if ($participantsSelectSupported) { "participants?`$select=$selectParticipants" } else { "participants" }
$allParticipants = @(Get-LandaxAll -endpoint $participantsEndpoint -headers $landaxHeaders)

foreach ($p in $allParticipants) {
if ($null -eq $p) { continue }

$linkTaskId = 0
foreach ($field in @("TaskId", "ObjectId", "ModuleId", "RecordId", "SourceId", "ParentId")) {
if ($p.PSObject.Properties.Name -contains $field -and $null -ne $p.$field) {
if ([int]::TryParse([string]$p.$field, [ref]$linkTaskId) -and $linkTaskId -gt 0) { break }
$linkTaskId = 0
}
}
if ($linkTaskId -eq 0) { continue }

$name = $null
if ($p.PSObject.Properties.Name -contains "CoworkerId" -and $null -ne $p.CoworkerId -and $coworkerById.ContainsKey([int]$p.CoworkerId)) {
$name = ([string]$coworkerById[[int]$p.CoworkerId]).Trim() -replace "\s+", " "
} elseif ($p.PSObject.Properties.Name -contains "Name" -and -not [string]::IsNullOrWhiteSpace([string]$p.Name)) {
$name = ([string]$p.Name).Trim() -replace "\s+", " "
}
if ([string]::IsNullOrWhiteSpace($name)) { continue }

if (-not $participantsByTaskId.ContainsKey($linkTaskId)) {
$participantsByTaskId[$linkTaskId] = New-Object System.Collections.Generic.List[string]
}
$participantsByTaskId[$linkTaskId].Add($name)
}

$participantsApiEnd = Get-Date
Write-StepTiming -Name "Participants API" -Start $participantsApiStart -End $participantsApiEnd
Write-Host "Participants hentet: $($allParticipants.Count)"
Write-Host "Unike TaskId: $($participantsByTaskId.Count)"

# =====================================================================
# TRINN 6 - Bygg Ressurser og produksjonstavle
# =====================================================================

$dashboardDataStart = Get-Date
$productionBoard = @(foreach ($t in $filteredTasks) {
$taskId        = [int]$t.Id
$resourceNames = New-Object System.Collections.Generic.List[string]
$seenNames     = @{}

if ($null -ne $t.HandledByCoworkerId -and $coworkerById.ContainsKey([int]$t.HandledByCoworkerId)) {
$name = ([string]$coworkerById[[int]$t.HandledByCoworkerId]).Trim() -replace "\s+", " "
$seenNames[$name] = $true
$resourceNames.Add($name)
}

if ($participantsByTaskId.ContainsKey($taskId)) {
foreach ($name in $participantsByTaskId[$taskId]) {
if (-not [string]::IsNullOrWhiteSpace($name) -and -not $seenNames.ContainsKey($name)) {
$seenNames[$name] = $true
$resourceNames.Add($name)
}
}
}

$ressurser = if ($resourceNames.Count -gt 0) { $resourceNames -join ", " } else { "IKKE TILDELT" }


[PSCustomObject]@{
Id             = $t.Id
Ressurser      = $ressurser
Oppgave        = $t.Oppgave
BoardStart     = $t.BoardStart
BoardEnd       = $t.BoardEnd
PlannedStartLocal = $t.PlannedStartLocal
PlannedDoneLocal  = $t.PlannedDoneLocal
Prosjekt       = $t.Prosjekt
ProsjektNummer = $t.ProsjektNummer
Lukket         = $t.Lukket
Status         = $t.Status
LandaxStatusRaw = $t.LandaxStatusRaw
}
})
$dashboardDataEnd = Get-Date
Write-StepTiming -Name "DashboardData" -Start $dashboardDataStart -End $dashboardDataEnd

$tasksWithoutResource = @($productionBoard | Where-Object { $_.Ressurser -eq "IKKE TILDELT" }).Count
$tasksHiddenFromWeeklyNoResource = $tasksWithoutResource
Write-Host ("Oppgaver uten ressurs: {0}" -f $tasksWithoutResource)
Write-Host ("Oppgaver skjult fra ukeplan fordi ressurs mangler: {0}" -f $tasksHiddenFromWeeklyNoResource)

# =====================================================================
# TRINN 7 - Bygg weeklyBoardData
# =====================================================================

function Get-BoardTaskText([object]$Task, [string]$StatusText) {
$prosjektLinje = if ($null -ne $Task.Prosjekt -and -not [string]::IsNullOrWhiteSpace([string]$Task.Prosjekt)) {
$pNum = if ($null -ne $Task.ProsjektNummer -and [string]$Task.ProsjektNummer -ne '') { "$($Task.ProsjektNummer) - " } else { "" }
"${pNum}$($Task.Prosjekt)"
} else { $null }

if ($null -ne $prosjektLinje) {
return "$prosjektLinje`n$($Task.Oppgave) ($StatusText)"
}
return "$($Task.Oppgave) ($StatusText)"
}

function Get-StatusCssClass([string]$StatusText) {
$normalized = if ([string]::IsNullOrWhiteSpace($StatusText)) { "" } else { $StatusText.Trim().ToLowerInvariant() }
switch ($normalized) {
"planlagt" { return "status-planlagt" }
"påbegynt" { return "status-pabegynt" }
"pabegynt" { return "status-pabegynt" }
"forsinket" { return "status-forsinket" }
"ferdig" { return "status-ferdig" }
default { return "status-landax" }
}
}

$weeklyBoardDataStart = Get-Date
$statusDiagnosticTaskId = 3427
$weeklyPersonDiagnostics = New-Object System.Collections.Generic.List[object]
$weeklyBoardData = @(
$productionBoard |
Where-Object { $_.Ressurser -ne "IKKE TILDELT" } |
ForEach-Object {
$task = $_
$persons = @($task.Ressurser -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ })
foreach ($person in $persons) {
[PSCustomObject]@{ Person = $person; Task = $task }
}
} |
Group-Object Person |
ForEach-Object {
$person = $_.Name
$tasksBeforeFilteringIds = @($_.Group | ForEach-Object {
if ($null -ne $_.Task -and $null -ne $_.Task.Id) { [int]$_.Task.Id }
} | Sort-Object -Unique)
$taskIdsAfterFilteringSet = New-Object 'System.Collections.Generic.HashSet[string]'
$byDay = @{
Mandag  = [System.Collections.Generic.List[object]]::new()
Tirsdag = [System.Collections.Generic.List[object]]::new()
Onsdag  = [System.Collections.Generic.List[object]]::new()
Torsdag = [System.Collections.Generic.List[object]]::new()
Fredag  = [System.Collections.Generic.List[object]]::new()
}

foreach ($row in $_.Group) {

$t  = $row.Task
$rs = $t.BoardStart
$re = $t.BoardEnd
if ($null -eq $rs -and $null -eq $re) { continue }
if ($null -eq $rs) { $rs = $re }
if ($null -eq $re) { $re = $rs }
if ($re -lt $rs) { $tmp = $rs; $rs = $re; $re = $tmp }

# Statusprioritet og debug
$baseStatus = [string]$t.Status
$landaxStatusRaw = if ($null -ne $t.LandaxStatusRaw) { [string]$t.LandaxStatusRaw } else { "" }
$normalizedLandaxStatus = if ([string]::IsNullOrWhiteSpace($landaxStatusRaw)) { "" } else { $landaxStatusRaw.Trim().ToLowerInvariant() }
$isLandaxNotStarted = ($normalizedLandaxStatus -eq "ikke påbegynt")
$isLandaxStarted = ($normalizedLandaxStatus -eq "påbegynt" -or $normalizedLandaxStatus -eq "pabegynt")
$isLandaxDone = ($normalizedLandaxStatus -in @("ferdig", "lukket", "closed", "completed"))
$plannedStartDiag = $t.PlannedStartLocal
$plannedDoneDiag = $t.PlannedDoneLocal
$doneDateDiag = $t.DoneDateLocal
$isTestTask = ([int]$t.Id -eq $statusDiagnosticTaskId -or [string]$t.Oppgave -eq "Test Oppgave")
$diagDates = New-Object System.Collections.Generic.List[string]
$diagStatuses = New-Object System.Collections.Generic.List[string]

# --- Forsinket-statuslogikk ---
$today = $todayOslo
$endDate = $t.PlannedDoneLocal
if ($null -eq $endDate) { $endDate = $t.BoardEnd }
$isDelayed = $false
if (-not $isLandaxDone -and $null -ne $endDate -and $endDate -lt $today) {
    $isDelayed = $true
}

$cursor = $rs
while ($cursor -lt $weekEnd -and $cursor -le $re) {
$dayName = switch ($cursor.DayOfWeek) {
"Monday"    { "Mandag" }
"Tuesday"   { "Tirsdag" }
"Wednesday" { "Onsdag" }
"Thursday"  { "Torsdag" }
"Friday"    { "Fredag" }
default     { $null }
}
if ($dayName -and $cursor -ge $weekStart) {
$cursorDateKey = $cursor.Date.ToString("yyyy-MM-dd")
$isHolidayDate = $holidayNameByDate.ContainsKey($cursorDateKey)
$isPlannedStartOnCursor = ($null -ne $t.PlannedStartLocal -and $t.PlannedStartLocal.Date -eq $cursor.Date)
$isPlannedDoneOnCursor = ($null -ne $t.PlannedDoneLocal -and $t.PlannedDoneLocal.Date -eq $cursor.Date)
if ($isHolidayDate -and -not ($isPlannedStartOnCursor -or $isPlannedDoneOnCursor)) {
$cursor = $cursor.AddDays(1)
continue
}


# --- Statusvisning med prioritet: Ferdig > Forsinket > Påbegynt > Planlagt ---
if ($isLandaxDone -or $t.Lukket -eq $true) {
    $renderedStatus = "Ferdig"
} elseif ($isDelayed) {
    $renderedStatus = "Forsinket"
} elseif ($isLandaxStarted -or $baseStatus -eq "Påbegynt") {
    $renderedStatus = "Påbegynt"
} elseif ($isLandaxNotStarted) {
    $renderedStatus = if ($cursor.Date -lt $todayOslo) { "Forsinket" } else { "Planlagt" }
} else {
    $renderedStatus = $baseStatus
}
$statusCssClass = Get-StatusCssClass -StatusText $renderedStatus
$taskText = Get-BoardTaskText -Task $t -StatusText $renderedStatus

# Debug-logg for statusvisning
Write-Host ("StatusVisDebug | TaskId: {0} | Status: {1} | EndDate: {2} | IsDelayed: {3} | RenderedStatus: {4}" -f [int]$t.Id, $baseStatus, $endDate, $isDelayed, $renderedStatus)

$byDay[$dayName].Add([PSCustomObject]@{
    Date = $cursor.Date
    TaskId = $t.Id
    Oppgave = $t.Oppgave
    Prosjekt = $t.Prosjekt
    Status = $renderedStatus
    StatusCssClass = $statusCssClass
    TaskText = $taskText
})
$null = $taskIdsAfterFilteringSet.Add([string]$t.Id)
}
$cursor = $cursor.AddDays(1)
}
if ($isTestTask -and $diagDates.Count -gt 0) {
Write-Host ("Statusdiagnostikk oppsummering -> TaskId: {0} | Person: {1} | Visningsdatoer: {2} | Status per visningsdato: {3}" -f [int]$t.Id, $person, ($diagDates -join ","), ($diagStatuses -join ","))
}
}

$taskIdsAfterFiltering = @($taskIdsAfterFilteringSet | ForEach-Object {
if (-not [string]::IsNullOrWhiteSpace($_)) { [int]$_ }
} | Sort-Object -Unique)
$visibleCellsInWeek = @(
"Mandag", "Tirsdag", "Onsdag", "Torsdag", "Fredag" |
ForEach-Object { if (@($byDay[$_]).Count -gt 0) { $_ } }
).Count
$inclusionReason = if ($visibleCellsInWeek -eq 0) {
"Personen blir lagt til i HTML-personlisten fordi vedkommende finnes i ressurslisten før dag/dato-filtrering. Group-Object Person bygger en rad, og raden beholdes selv om alle dagceller blir tomme etter filtrering."
} else {
"Personen blir lagt til i HTML-personlisten fordi vedkommende har minst en synlig celle etter filtrering i ukeplanen."
}
$weeklyPersonDiagnostics.Add([PSCustomObject]@{
PersonNavn = $person
TasksBeforeFilteringIds = @($tasksBeforeFilteringIds)
TasksAfterFilteringIds = @($taskIdsAfterFiltering)
VisibleCellsInWeek = $visibleCellsInWeek
InclusionReason = $inclusionReason
}) | Out-Null

[PSCustomObject]@{
Person  = $person
DayItems = @{
Mandag  = @($byDay["Mandag"])
Tirsdag = @($byDay["Tirsdag"])
Onsdag  = @($byDay["Onsdag"])
Torsdag = @($byDay["Torsdag"])
Fredag  = @($byDay["Fredag"])
}
}
} |
Sort-Object Person
)
$weeklyBoardDataEnd = Get-Date
Write-StepTiming -Name "WeeklyBoardData" -Start $weeklyBoardDataStart -End $weeklyBoardDataEnd

$targetPersonName = "Giannis Vamvouras"
$targetPersonDiagnostics = @($weeklyPersonDiagnostics | Where-Object { [string]$_.PersonNavn -eq $targetPersonName })
if ($targetPersonDiagnostics.Count -eq 0) {
Write-Host ("WeeklyBoardData diagnostikk: Fant ikke person '{0}' i personlisten denne kjøringen." -f $targetPersonName)
} else {
foreach ($diag in $targetPersonDiagnostics) {
Write-Host ("PersonNavn: {0}" -f [string]$diag.PersonNavn)
Write-Host ("Antall oppgaver før filtrering: {0}" -f @($diag.TasksBeforeFilteringIds).Count)
Write-Host ("Antall oppgaver etter filtrering: {0}" -f @($diag.TasksAfterFilteringIds).Count)
Write-Host ("TaskId-er etter filtrering: {0}" -f ((@($diag.TasksAfterFilteringIds) -join ", ")))
Write-Host ("Antall synlige celler i ukeplanen: {0}" -f [int]$diag.VisibleCellsInWeek)
Write-Host ("Hvorfor lagt til i personlisten: {0}" -f [string]$diag.InclusionReason)
}
}

# =====================================================================
# TRINN 8 - Generer HTML
# =====================================================================

function Convert-StatusTokenToHtml([string]$Token) {
$trimmed = if ($null -eq $Token) { "" } else { $Token.Trim() }
switch ($trimmed) {
"Planlagt"  { return "<span class='status status-planlagt'>Planlagt</span>" }
"Påbegynt"  { return "<span class='status status-pabegynt'>P&aring;begynt</span>" }
"Pabegynt"  { return "<span class='status status-pabegynt'>P&aring;begynt</span>" }
"Pagar"     { return "<span class='status status-pabegynt'>P&aring;begynt</span>" }
"Pågår"     { return "<span class='status status-pabegynt'>P&aring;begynt</span>" }
"Forsinket" { return "<span class='status status-forsinket'>Forsinket</span>" }
"Ferdig"    { return "<span class='status status-ferdig'>Ferdig</span>" }
default      { return "<span class='status status-landax'>$([System.Net.WebUtility]::HtmlEncode($trimmed))</span>" }
}
}

function Convert-WeeklyCellToHtml([string]$Text) {
if ([string]::IsNullOrWhiteSpace($Text)) { return "<div class='empty'>-</div>" }
$lines = @($Text -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
if ($lines.Count -eq 0) { return "<div class='empty'>-</div>" }
$statusPattern = '\(([^)]*)\)$'

function Convert-TaskLineToHtml([string]$TaskLine, [string]$Pattern) {
 if ($TaskLine -match $Pattern) {
  $statusText = $matches[1]
  $prefixText = ($TaskLine -replace '\s*\([^)]*\)$', '')
  $prefixHtml = [System.Net.WebUtility]::HtmlEncode($prefixText)
  $statusParts = @($statusText -split '\s*·\s*' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
  $statusHtml = @($statusParts | ForEach-Object { Convert-StatusTokenToHtml -Token $_ }) -join "<span class='status-sep'> &middot; </span>"
  return "$prefixHtml (<span class='status-group'>$statusHtml</span>)"
 }
 return [System.Net.WebUtility]::HtmlEncode($TaskLine)
}

$projectOrder = New-Object System.Collections.Generic.List[string]
$tasksByProject = @{}

for ($i = 0; $i -lt $lines.Count; $i++) {
 $line = $lines[$i]
 $lineIsTask = ($line -match $statusPattern)
 $nextIsTask = ($i + 1 -lt $lines.Count -and $lines[$i + 1] -match $statusPattern)

 $projectKey = ""
 $taskLine = $line

 if (-not $lineIsTask -and $nextIsTask) {
  $projectKey = $line
  $taskLine = $lines[$i + 1]
  $i++
 }

 if (-not $tasksByProject.ContainsKey($projectKey)) {
  $tasksByProject[$projectKey] = New-Object System.Collections.Generic.List[string]
  $projectOrder.Add($projectKey) | Out-Null
 }
 $tasksByProject[$projectKey].Add($taskLine) | Out-Null
}

$groupsHtml = foreach ($project in $projectOrder) {
 $projectTitleHtml = ""
 if (-not [string]::IsNullOrWhiteSpace($project)) {
  $projectTitleHtml = "<div class='project-title'>$([System.Net.WebUtility]::HtmlEncode($project))</div>"
 }

 $taskItems = foreach ($taskLine in $tasksByProject[$project]) {
  $taskHtml = Convert-TaskLineToHtml -TaskLine $taskLine -Pattern $statusPattern
  "<div class='task-item'>&bull; $taskHtml</div>"
 }

 "<div class='project-group'>$projectTitleHtml$($taskItems -join '')</div>"
}

return ($groupsHtml -join "")
}

$calendar   = [System.Globalization.CultureInfo]::GetCultureInfo("nb-NO").Calendar
$weekRule   = [System.Globalization.CalendarWeekRule]::FirstFourDayWeek
$firstDow   = [System.DayOfWeek]::Monday
$weekNumber = $calendar.GetWeekOfYear((Get-Date), $weekRule, $firstDow)
$todayLabel = (Get-Date).ToString("dddd dd.MM.yyyy", [System.Globalization.CultureInfo]::GetCultureInfo("nb-NO"))
$lastUpdated = Get-Date -Format "HH:mm"

$duplicateGroupsBeforeHtml = 0
$duplicatesRemovedBeforeHtml = 0
$remainingTaskIdDuplicatesAfterHtmlDedupe = 0

function Get-FinalRenderDedupeKey([string]$Person, [string]$DayName, [object]$Item) {
if ($null -ne $Item.TaskId -and -not [string]::IsNullOrWhiteSpace([string]$Item.TaskId)) {
return "$Person|$DayName|$([string]$Item.TaskId)"
}
return "$Person|$DayName|$([string]$Item.Oppgave)|$([string]$Item.Prosjekt)|$([string]$Item.Status)"
}

$weekdayDateByName = @{
Mandag  = $weekStart.Date
Tirsdag = $weekStart.AddDays(1).Date
Onsdag  = $weekStart.AddDays(2).Date
Torsdag = $weekStart.AddDays(3).Date
Fredag  = $weekStart.AddDays(4).Date
}

function Get-HolidayColumnClass([datetime]$DateValue) {
$dateKey = $DateValue.Date.ToString("yyyy-MM-dd")
if ($holidayNameByDate.ContainsKey($dateKey)) {
return "holiday-column"
}
return ""
}

$mandagHolidayClass = Get-HolidayColumnClass -DateValue $weekdayDateByName["Mandag"]
$tirsdagHolidayClass = Get-HolidayColumnClass -DateValue $weekdayDateByName["Tirsdag"]
$onsdagHolidayClass = Get-HolidayColumnClass -DateValue $weekdayDateByName["Onsdag"]
$torsdagHolidayClass = Get-HolidayColumnClass -DateValue $weekdayDateByName["Torsdag"]
$fredagHolidayClass = Get-HolidayColumnClass -DateValue $weekdayDateByName["Fredag"]

$mandagHeaderClassAttr = if ([string]::IsNullOrWhiteSpace($mandagHolidayClass)) { "" } else { " class=`"$mandagHolidayClass`"" }
$tirsdagHeaderClassAttr = if ([string]::IsNullOrWhiteSpace($tirsdagHolidayClass)) { "" } else { " class=`"$tirsdagHolidayClass`"" }
$onsdagHeaderClassAttr = if ([string]::IsNullOrWhiteSpace($onsdagHolidayClass)) { "" } else { " class=`"$onsdagHolidayClass`"" }
$torsdagHeaderClassAttr = if ([string]::IsNullOrWhiteSpace($torsdagHolidayClass)) { "" } else { " class=`"$torsdagHolidayClass`"" }
$fredagHeaderClassAttr = if ([string]::IsNullOrWhiteSpace($fredagHolidayClass)) { "" } else { " class=`"$fredagHolidayClass`"" }

$mandagCellClassAttr = if ([string]::IsNullOrWhiteSpace($mandagHolidayClass)) { "" } else { " class=`"$mandagHolidayClass`"" }
$tirsdagCellClassAttr = if ([string]::IsNullOrWhiteSpace($tirsdagHolidayClass)) { "" } else { " class=`"$tirsdagHolidayClass`"" }
$onsdagCellClassAttr = if ([string]::IsNullOrWhiteSpace($onsdagHolidayClass)) { "" } else { " class=`"$onsdagHolidayClass`"" }
$torsdagCellClassAttr = if ([string]::IsNullOrWhiteSpace($torsdagHolidayClass)) { "" } else { " class=`"$torsdagHolidayClass`"" }
$fredagCellClassAttr = if ([string]::IsNullOrWhiteSpace($fredagHolidayClass)) { "" } else { " class=`"$fredagHolidayClass`"" }

$weeklyBoardDataTotalPersons = @($weeklyBoardData).Count
$weeklyBoardDataFilteredForHtml = @(
$weeklyBoardData | Where-Object {
(
@($_.DayItems["Mandag"]).Count +
@($_.DayItems["Tirsdag"]).Count +
@($_.DayItems["Onsdag"]).Count +
@($_.DayItems["Torsdag"]).Count +
@($_.DayItems["Fredag"]).Count
) -gt 0
}
)
$removedPersonsWithoutVisibleCells = $weeklyBoardDataTotalPersons - @($weeklyBoardDataFilteredForHtml).Count
Write-Host ("Personer fjernet før HTML (0 synlige celler): {0}" -f $removedPersonsWithoutVisibleCells)

$weeklyBoardDataForRender = @(
$weeklyBoardDataFilteredForHtml | ForEach-Object {
$person = [string]$_.Person
$renderByDay = @{}

foreach ($dayName in @("Mandag", "Tirsdag", "Onsdag", "Torsdag", "Fredag")) {
$entries = @($_.DayItems[$dayName])
$counts = @{}
foreach ($entry in $entries) {
$key = Get-FinalRenderDedupeKey -Person $person -DayName $dayName -Item $entry
if (-not $counts.ContainsKey($key)) { $counts[$key] = 0 }
$counts[$key]++
}

$seenKeys = New-Object 'System.Collections.Generic.HashSet[string]'
$dedupedItems = New-Object System.Collections.Generic.List[object]
foreach ($entry in $entries) {
$key = Get-FinalRenderDedupeKey -Person $person -DayName $dayName -Item $entry
if ($seenKeys.Add($key)) {
$dedupedItems.Add($entry) | Out-Null
} elseif ($counts[$key] -gt 1) {
$duplicateGroupsBeforeHtml++
$duplicatesRemovedBeforeHtml++
$taskIdForLog = if ($null -ne $entry.TaskId -and -not [string]::IsNullOrWhiteSpace([string]$entry.TaskId)) { [string]$entry.TaskId } else { "(mangler)" }
Write-Host ("Duplikat -> Person: {0} | Ukedag: {1} | TaskId: {2} | Oppgave: {3} | Antall før: {4} | Antall etter: 1" -f $person, $dayName, $taskIdForLog, [string]$entry.Oppgave, $counts[$key])
}
}

$remainingTaskIdDuplicatesAfterHtmlDedupe += @($dedupedItems | Where-Object { $null -ne $_.TaskId -and -not [string]::IsNullOrWhiteSpace([string]$_.TaskId) } | Group-Object -Property { [string]$_.TaskId } | Where-Object { $_.Count -gt 1 }).Count

$renderByDay[$dayName] = @($dedupedItems | ForEach-Object { $_.TaskText })
}

[PSCustomObject]@{
Person  = $person
Mandag  = ($renderByDay["Mandag"]  -join "`n")
Tirsdag = ($renderByDay["Tirsdag"] -join "`n")
Onsdag  = ($renderByDay["Onsdag"]  -join "`n")
Torsdag = ($renderByDay["Torsdag"] -join "`n")
Fredag  = ($renderByDay["Fredag"]  -join "`n")
}
}
)

Write-Host ("Duplikater før HTML: {0}" -f $duplicateGroupsBeforeHtml)
Write-Host ("Duplikater fjernet før HTML: {0}" -f $duplicatesRemovedBeforeHtml)
Write-Host ("Gjenstaende Person+Ukedag+TaskId-duplikater etter dedupe: {0}" -f $remainingTaskIdDuplicatesAfterHtmlDedupe)

# Standard dagcelle-rendering uten colspan
$tableRows = ($weeklyBoardDataForRender | ForEach-Object {
$person  = [System.Net.WebUtility]::HtmlEncode([string]$_.Person)
$mandag  = Convert-WeeklyCellToHtml -Text ([string]$_.Mandag)
$tirsdag = Convert-WeeklyCellToHtml -Text ([string]$_.Tirsdag)
$onsdag  = Convert-WeeklyCellToHtml -Text ([string]$_.Onsdag)
$torsdag = Convert-WeeklyCellToHtml -Text ([string]$_.Torsdag)
$fredag  = Convert-WeeklyCellToHtml -Text ([string]$_.Fredag)
@"
<tr>
<td class="assignee">$person</td>
<td$mandagCellClassAttr>$mandag</td>
<td$tirsdagCellClassAttr>$tirsdag</td>
<td$onsdagCellClassAttr>$onsdag</td>
<td$torsdagCellClassAttr>$torsdag</td>
<td$fredagCellClassAttr>$fredag</td>
</tr>
"@
}) -join "`n"

$htmlGenerationStart = Get-Date
$html = @"
<!doctype html>
<html lang="nb">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<meta http-equiv="refresh" content="$($script:RefreshIntervalSeconds)" />
<title>Produksjonstavle</title>
<style>
:root {
--bg: #0f1115;
--panel: #171a21;
--grid: #2a2f3a;
--text: #f4f6fb;
--muted: #9ea7b8;
--planlagt: #9ca3af;
--pabegynt: #4da3ff;
--forsinket: #ff5b5b;
--ferdig: #47c97d;
}

* { box-sizing: border-box; }

html, body {
margin: 0;
padding: 0;
width: 100%;
height: 100%;
background: var(--bg);
color: var(--text);
font-family: "Segoe UI", "Helvetica Neue", sans-serif;
}

.screen {
min-height: 100vh;
padding: 1.2rem;
display: flex;
flex-direction: column;
gap: 1rem;
}

.header {
display: flex;
justify-content: space-between;
align-items: baseline;
gap: 1rem;
flex-wrap: wrap;
}

.title {
font-size: clamp(1.8rem, 2.6vw, 3rem);
font-weight: 800;
letter-spacing: 0.02em;
}

.meta {
font-size: clamp(1rem, 1.3vw, 1.4rem);
color: var(--muted);
font-weight: 600;
}

.board {
flex: 1;
background: var(--panel);
border: 1px solid var(--grid);
border-radius: 12px;
overflow: hidden;
}

table {
width: 100%;
border-collapse: collapse;
table-layout: fixed;
}

th, td {
border: 1px solid var(--grid);
vertical-align: top;
padding: 0.65rem 0.55rem;
font-size: clamp(0.88rem, 1.03vw, 1.15rem);
line-height: 1.35;
white-space: pre-wrap;
word-break: break-word;
}

tbody tr {
height: auto;
}

th {
position: sticky;
top: 0;
z-index: 2;
background: #1e2430;
text-align: left;
font-size: clamp(0.95rem, 1.15vw, 1.25rem);
font-weight: 800;
letter-spacing: 0.01em;
}

.assignee {
width: 18ch;
font-weight: 700;
background: #151a23;
position: sticky;
left: 0;
z-index: 1;
}

.task-line { margin-bottom: 0.28rem; }
.task-line:last-child { margin-bottom: 0; }
.project-line {
font-size: 0.92em;
color: var(--muted);
font-weight: 500;
}
.task-title {
font-weight: 700;
color: var(--text);
}
.empty { color: #6f7685; }

.status { font-weight: 700; }
.status-planlagt { color: var(--planlagt); }
.status-pabegynt { color: var(--pabegynt); }
.status-forsinket { color: var(--forsinket); }
.status-ferdig { color: var(--ferdig); }
.status-landax { color: #d8deea; }
.status-sep { color: var(--muted); }

.holiday-column {
background: #3a2c2c;
}

.project-group {
margin-bottom: 6px;
}

.project-group:last-child {
margin-bottom: 0;
}

.project-title {
font-size: 0.85em;
opacity: .75;
font-weight: 600;
}

.task-item {
margin-left: 10px;
margin-bottom: 2px;
}

body.compact .screen {
padding: 0.95rem;
gap: 0.75rem;
}

body.compact .title {
font-size: clamp(1.55rem, 2.2vw, 2.4rem);
}

body.compact .meta {
font-size: clamp(0.92rem, 1.15vw, 1.15rem);
}

body.compact th,
body.compact td {
padding: 0.5rem 0.4rem;
font-size: clamp(0.8rem, 0.94vw, 1.02rem);
line-height: 1.28;
}

body.compact .task-line {
margin-bottom: 0.2rem;
}

body.dense .screen {
padding: 0.7rem;
gap: 0.55rem;
}

body.dense .title {
font-size: clamp(1.35rem, 1.9vw, 2.05rem);
}

body.dense .meta {
font-size: clamp(0.82rem, 0.98vw, 1rem);
}

body.dense th,
body.dense td {
padding: 0.34rem 0.28rem;
font-size: clamp(0.72rem, 0.84vw, 0.92rem);
line-height: 1.2;
}

body.dense .task-line {
margin-bottom: 0.12rem;
}

body.dense .project-line {
font-size: 0.88em;
}

@media (max-width: 1200px) {
.screen { padding: 0.8rem; }
th, td { padding: 0.45rem; }
.assignee { width: 14ch; }
}
</style>
</head>
<body>
<div class="screen">
<div class="header">
<div class="title">Produksjonstavle Ukeplan</div>
<div class="meta">$todayLabel &middot; Uke $weekNumber &middot; Sist oppdatert: $lastUpdated<span id="page-indicator" style="display:none"></span></div>
</div>

<div class="board">
<table>
<thead>
<tr>
<th>Person</th>
<th$mandagHeaderClassAttr>Mandag</th>
<th$tirsdagHeaderClassAttr>Tirsdag</th>
<th$onsdagHeaderClassAttr>Onsdag</th>
<th$torsdagHeaderClassAttr>Torsdag</th>
<th$fredagHeaderClassAttr>Fredag</th>
</tr>
</thead>
<tbody>
$tableRows
</tbody>
</table>
</div>
</div>
<script>
(function () {
    const SCALE_CLASSES = ["compact", "dense"];
    const PAGE_INTERVAL_MS = 12000;
    const MIN_ROW_HEIGHT = 55;
    const MAX_ROW_HEIGHT = 160;

    let currentPage = 0;
    let pages = [];
    let pageTimer = null;

    function setMode(mode) {
        document.body.classList.remove(...SCALE_CLASSES);
        if (mode === "compact" || mode === "dense") {
            document.body.classList.add(mode);
        }
    }

    function fitsViewport() {
        const screen = document.querySelector(".screen");
        if (!screen) return true;
        return Math.ceil(screen.getBoundingClientRect().height) <= window.innerHeight;
    }

    function computeAvailableHeight() {
        const screen = document.querySelector(".screen");
        const header = document.querySelector(".header");
        const board = document.querySelector(".board");
        const tableHead = document.querySelector("thead");
        const viewportHeight = window.innerHeight;
        const screenStyles = screen ? window.getComputedStyle(screen) : null;
        const paddingY = screenStyles
            ? (parseFloat(screenStyles.paddingTop) || 0) + (parseFloat(screenStyles.paddingBottom) || 0)
            : 0;
        const headerHeight = header ? Math.ceil(header.getBoundingClientRect().height) : 0;
        const tableHeadHeight = tableHead ? Math.ceil(tableHead.getBoundingClientRect().height) : 0;
        const boardPaddingY = board
            ? ((parseFloat(window.getComputedStyle(board).paddingTop) || 0) + (parseFloat(window.getComputedStyle(board).paddingBottom) || 0))
            : 0;
        let available = viewportHeight - headerHeight - tableHeadHeight - paddingY - boardPaddingY;
        if (available <= 0 && board) {
            available = board.clientHeight - tableHeadHeight;
        }
        return available;
    }

    function buildPages(allRows, rowsPerPage) {
        const result = [];
        for (let i = 0; i < allRows.length; i += rowsPerPage) {
            result.push(allRows.slice(i, i + rowsPerPage));
        }
        return result;
    }

    function updateIndicator(current, total) {
        const indicator = document.getElementById("page-indicator");
        if (!indicator) return;
        if (total <= 1) {
            indicator.textContent = "";
            indicator.style.display = "none";
        } else {
            indicator.textContent = " \u00b7 Side " + current + "/" + total;
            indicator.style.display = "";
        }
    }

    function showPage(pageIndex) {
        if (pages.length === 0) return;
        const allRows = document.querySelectorAll("tbody tr");
        allRows.forEach(function (r) { r.style.display = "none"; });
        const page = pages[pageIndex];
        page.forEach(function (r) { r.style.display = ""; });
        updateIndicator(pageIndex + 1, pages.length);
        console.debug("Pagination: side", pageIndex + 1, "av", pages.length);
    }

    function startPageTimer() {
        if (pageTimer) clearInterval(pageTimer);
        if (pages.length <= 1) return;
        pageTimer = setInterval(function () {
            currentPage = (currentPage + 1) % pages.length;
            showPage(currentPage);
        }, PAGE_INTERVAL_MS);
    }

    function update() {
        // 1. Apply best scale mode
        let mode = "normal";
        setMode("normal");
        if (!fitsViewport()) {
            mode = "compact";
            setMode("compact");
            if (!fitsViewport()) {
                mode = "dense";
                setMode("dense");
            }
        }
        console.debug("Board scale mode:", mode);

        // 2. Gather all rows and reset visibility
        const allRows = Array.from(document.querySelectorAll("tbody tr"));
        if (allRows.length === 0) return;
        allRows.forEach(function (r) { r.style.display = ""; r.style.height = ""; });

        // 3. Compute rows-per-page: max 7 persons per page
        const MAX_ROWS_PER_PAGE = 7;
        const available = computeAvailableHeight();
        const rowsPerPage = Math.min(MAX_ROWS_PER_PAGE, allRows.length);
        const rowHeight = Math.max(MIN_ROW_HEIGHT, Math.min(MAX_ROW_HEIGHT, available / rowsPerPage));

        // 4. Set height on ALL rows so hidden rows maintain size when shown
        allRows.forEach(function (r) { r.style.height = rowHeight + "px"; });

        // 5. Build pages
        pages = buildPages(allRows, rowsPerPage);
        currentPage = 0;

        // 6. Show first page and start timer
        showPage(currentPage);
        startPageTimer();

        console.debug("Rows:", allRows.length);
        console.debug("Rows per page:", rowsPerPage);
        console.debug("Pages:", pages.length);
    }

    window.addEventListener("load", update);
    window.addEventListener("resize", function () {
        window.requestAnimationFrame(update);
    });
})();
</script>
</body>
</html>
"@

$htmlPath = $HtmlOutputPath
$html | Out-File -Path $htmlPath -Encoding utf8
$htmlGenerationEnd = Get-Date
Write-StepTiming -Name "HTML-generering" -Start $htmlGenerationStart -End $htmlGenerationEnd

Write-Host "Publisering til SharePoint: Start"
try {
if (-not (Test-Path $HtmlOutputPath)) {
Write-Warning "Publisering til SharePoint feilet: HTML-fil ikke funnet"
} elseif (-not (Test-Path $SharePointPublishPath)) {
Write-Warning "Publisering til SharePoint feilet: mappe ikke funnet"
} else {
Copy-Item `
    $HtmlOutputPath `
    (Join-Path $SharePointPublishPath "production-board.html") `
    -Force
Write-Host "Publisering til SharePoint: OK"
}
} catch {
Write-Warning ("Publisering til SharePoint feilet: " + $_.Exception.Message)
}

# --- Snapshot og endringssporing ---

$currentTaskIds      = @($productionBoard | ForEach-Object { [int]$_.Id })
$currentResourceMap  = @{}
foreach ($row in $productionBoard) { $currentResourceMap[[string]$row.Id] = $row.Ressurser }

$participantNameSet = @{}
foreach ($names in $participantsByTaskId.Values) { foreach ($n in $names) { $participantNameSet[$n] = $true } }
$currentParticipantNames = @($participantNameSet.Keys | Sort-Object -Unique)

$previousTaskIds     = @()
$previousResourceMap = @{}
if ($null -ne $previousSnapshot) {
if ($previousSnapshot.PSObject.Properties.Name -contains "taskIds") {
$previousTaskIds = @($previousSnapshot.taskIds | ForEach-Object { [int]$_ })
}
if ($previousSnapshot.PSObject.Properties.Name -contains "taskResourceMap") {
foreach ($prop in $previousSnapshot.taskResourceMap.PSObject.Properties) {
$previousResourceMap[$prop.Name] = [string]$prop.Value
}
}
}

$newTasks = @($currentTaskIds | Where-Object { $_ -notin $previousTaskIds })


$changedRows = @($productionBoard | Where-Object {
$key = [string]$_.Id
(-not $previousResourceMap.ContainsKey($key)) -or ($previousResourceMap[$key] -ne $_.Ressurser)
})


$snapshotPayload = [PSCustomObject]@{
generatedAt      = (Get-Date).ToString("o")
taskIds          = @($currentTaskIds | Sort-Object -Unique)
participantNames = $currentParticipantNames
taskResourceMap  = $currentResourceMap
}
$snapshotPayload | ConvertTo-Json -Depth 20 | Out-File -Path $snapshotPath -Encoding utf8

# --- Kjøretid og API-statistikk ---

$elapsed       = [Math]::Round(((Get-Date) - $scriptStart).TotalSeconds, 1)
$totalApiCalls = $script:apiCallCount

$perfPayload = [PSCustomObject]@{
generatedAt = (Get-Date).ToString("o")
seconds = $elapsed
apiCalls = $totalApiCalls
}
$perfPayload | ConvertTo-Json -Depth 5 | Out-File -Path $perfPath -Encoding utf8

Write-Host "Total kjøretid: ${elapsed}s"