$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectFile = Join-Path $root "projects.xlsx"
$updateScript = Join-Path $root "update_projects.py"

$python = Get-Command python -ErrorAction SilentlyContinue
$pythonArgs = @()
if (-not $python) {
    $python = Get-Command py -ErrorAction SilentlyContinue
    $pythonArgs = @("-3")
}
if (-not $python) {
    Write-Error "Python was not found. Please install Python or run update_projects.py from Codex."
    exit 1
}

function Invoke-ProjectUpdate {
    & $python.Source @pythonArgs $updateScript
}

Write-Host "Watching projects.xlsx. Press Ctrl+C to stop."
Invoke-ProjectUpdate

$lastWrite = (Get-Item $projectFile).LastWriteTimeUtc
while ($true) {
    Start-Sleep -Seconds 2
    $currentWrite = (Get-Item $projectFile).LastWriteTimeUtc
    if ($currentWrite -ne $lastWrite) {
        $lastWrite = $currentWrite
        Start-Sleep -Seconds 1
        Invoke-ProjectUpdate
    }
}
