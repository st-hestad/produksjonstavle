# Finn rotmappe basert på hvor denne filen ligger
$ScriptRoot = $PSScriptRoot
$ProjectRoot = Resolve-Path (Join-Path $ScriptRoot "..")

# Forventer at LANDAX_BASEURL, LANDAX_USERNAME og LANDAX_PASSWORD allerede er satt som miljøvariabler
foreach ($var in @("LANDAX_BASEURL", "LANDAX_USERNAME", "LANDAX_PASSWORD")) {
    if (-not (Get-Item "env:$var" -ErrorAction SilentlyContinue)) {
        Write-Error "Miljøvariabel '$var' er ikke satt. Avbryter."
        exit 1
    }
}

# Kjør produksjonstavlen
& (Join-Path $ScriptRoot "landax-production-board.ps1")

$sourceHtml = Join-Path $ProjectRoot "production-board.html"
$targetHtml = Join-Path $ProjectRoot "index.html"
Copy-Item -Path $sourceHtml -Destination $targetHtml -Force
Write-Host "Kopierte production-board.html til $targetHtml"
