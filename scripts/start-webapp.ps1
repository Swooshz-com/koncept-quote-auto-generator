param(
  [string]$BindHost = "127.0.0.1",
  [int]$Port = 8765,
  [string]$PythonPath = ""
)

$ErrorActionPreference = "Stop"

function Find-Python {
  param([string]$RequestedPath)

  if ($RequestedPath -and (Test-Path -LiteralPath $RequestedPath)) {
    return (Resolve-Path -LiteralPath $RequestedPath).Path
  }

  if ($env:PYTHON -and (Test-Path -LiteralPath $env:PYTHON)) {
    return (Resolve-Path -LiteralPath $env:PYTHON).Path
  }

  $codexPython = Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
  if (Test-Path -LiteralPath $codexPython) {
    return $codexPython
  }

  $python = Get-Command python.exe -ErrorAction SilentlyContinue | Select-Object -First 1
  if ($python -and $python.Source -notmatch "\\WindowsApps\\") { return $python.Source }

  $py = Get-Command py.exe -ErrorAction SilentlyContinue | Select-Object -First 1
  if ($py -and $py.Source -notmatch "\\WindowsApps\\") { return $py.Source }

  throw "No Python runtime found. Pass -PythonPath C:\path\to\python.exe."
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$pythonExe = Find-Python -RequestedPath $PythonPath
$url = "http://${BindHost}:${Port}/"
$healthUrl = "http://${BindHost}:${Port}/api/health"
$logRoot = Join-Path $repoRoot "_logs"
$serverLogRoot = Join-Path $logRoot "server"
$logPath = Join-Path $serverLogRoot "webapp-server-$Port.out.log"
$errPath = Join-Path $serverLogRoot "webapp-server-$Port.err.log"

New-Item -ItemType Directory -Force -Path $serverLogRoot | Out-Null

# Some Windows agent shells expose both Path and PATH. Start-Process can fail
# when inheriting that duplicate environment dictionary.
$processEnv = [Environment]::GetEnvironmentVariables("Process")
if ($processEnv.Contains("Path") -and $processEnv.Contains("PATH")) {
  [Environment]::SetEnvironmentVariable("PATH", $null, "Process")
}

$existing = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1
if ($existing) {
  try {
    Invoke-WebRequest -Uri $healthUrl -UseBasicParsing -TimeoutSec 5 | Out-Null
    Write-Host "Swooshz Quote Generator already running at $url"
    exit 0
  } catch {
    throw "Port $Port is already in use, but $healthUrl did not respond cleanly. Stop that process or choose another -Port."
  }
}

Remove-Item -LiteralPath $logPath, $errPath -ErrorAction SilentlyContinue

$process = Start-Process -FilePath $pythonExe `
  -ArgumentList @("webapp/server.py", "--host", $BindHost, "--port", "$Port") `
  -WorkingDirectory $repoRoot `
  -RedirectStandardOutput $logPath `
  -RedirectStandardError $errPath `
  -WindowStyle Hidden `
  -PassThru

$deadline = (Get-Date).AddSeconds(15)
do {
  Start-Sleep -Milliseconds 400
  try {
    Invoke-WebRequest -Uri $healthUrl -UseBasicParsing -TimeoutSec 2 | Out-Null
    Write-Host "Swooshz Quote Generator running at $url"
    Write-Host "Process: $($process.Id)"
    Write-Host "Logs: $logPath"
    Write-Host "Errors: $errPath"
    exit 0
  } catch {
    if ($process.HasExited) {
      $stderr = if (Test-Path -LiteralPath $errPath) { Get-Content -Raw -Path $errPath } else { "" }
      throw "Server exited before becoming healthy. $stderr"
    }
  }
} while ((Get-Date) -lt $deadline)

throw "Server did not become healthy within 15 seconds. Check $logPath and $errPath."
