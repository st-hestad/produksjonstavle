$script:LandaxBaseUrlDefault = "https://mekon.landax.no/api/v32"

$script:RefreshIntervalSeconds = 300

$script:ProductionBoardLogDir = Join-Path $PSScriptRoot "../04_logs"
$script:ProductionBoardLogPath = Join-Path $script:ProductionBoardLogDir "production-board.log"
$script:SnapshotPath = Join-Path $script:ProductionBoardLogDir "production_board_snapshot.json"
$script:PerfPath = Join-Path $script:ProductionBoardLogDir "production_board_perf.json"
$script:BaselinePerfPath = Join-Path $script:ProductionBoardLogDir "production_board_perf_baseline.json"

$script:DepartmentsCachePath = Join-Path $script:ProductionBoardLogDir "cache_departments.json"
$script:ProjectsCachePath = Join-Path $script:ProductionBoardLogDir "cache_projects.json"
$script:CoworkersCachePath = Join-Path $script:ProductionBoardLogDir "cache_coworkers.json"

$script:HtmlOutputPath = Join-Path $PSScriptRoot "../production-board.html"

$script:DefaultFullScan = $false
$script:DefaultOptimizedApi = $false
$script:DefaultApiMode = "FastDefault"
