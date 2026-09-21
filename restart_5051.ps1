# restart_5051.ps1 - restart the :5051 server on the current source.  ASCII ONLY.
#
# What it does (measured 2026-09-21):
#   1. finds the process listening on :5051 and, if its parent is a
#      start_server.bat launcher loop, stops that loop FIRST (otherwise the
#      old loop restarts the old code 5 s later and two servers fight for
#      the port - serve.py can bind twice on Windows without an error).
#   2. stops the :5051 holder.
#   3. runs start_server.bat (frees the port again, runs serve.py in a loop).
#   4. waits for the new listener and GETs /login.
# If the old server was started from an elevated shell (the ChatGPT deploy
# scripts did that), stopping it is denied; then this script relaunches
# itself elevated (UAC prompt) and does the same steps there.
# The :5052 domain server is never touched.

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $MyInvocation.MyCommand.Path
$me = $MyInvocation.MyCommand.Path
$port = 5051

function Port-Pids([int]$number) {
    @(netstat -ano -p tcp | ForEach-Object {
        if ($_ -match "^\s*TCP\s+\S+:$number\s+\S+\s+LISTENING\s+(\d+)\s*$") { [int]$Matches[1] }
    } | Sort-Object -Unique)
}

function Stop-Server {
    # returns $true when everything that held the port is gone
    $ok = $true
    foreach ($p in (Port-Pids $port)) {
        $proc = Get-CimInstance Win32_Process -Filter "ProcessId = $p" -ErrorAction SilentlyContinue
        if ($proc) {
            $parent = Get-CimInstance Win32_Process -Filter ("ProcessId = " + $proc.ParentProcessId) -ErrorAction SilentlyContinue
            if ($parent -and $parent.Name -eq 'cmd.exe' -and $parent.CommandLine -like '*start_server.bat*') {
                Log ("stopping old launcher loop cmd pid " + $parent.ProcessId)
                try { Stop-Process -Id $parent.ProcessId -Force -ErrorAction Stop } catch { Log ("  denied: " + $_.Exception.Message); $ok = $false }
            }
        }
        Log ("stopping :" + $port + " holder pid " + $p)
        try { Stop-Process -Id $p -Force -ErrorAction Stop } catch { Log ("  denied: " + $_.Exception.Message); $ok = $false }
    }
    Start-Sleep -Seconds 2
    if ((Port-Pids $port).Count -gt 0) { $ok = $false }
    return $ok
}

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$log = Join-Path $repo 'restart_5051.log'
function Log([string]$m) { Write-Host $m; Add-Content -Path $log -Value ((Get-Date).ToString('s') + ' ' + $m) }
Log ("ps1 start  repo: " + $repo + "  elevated: " + $isAdmin + "  holders: " + ((Port-Pids $port) -join ','))

if (-not (Stop-Server)) {
    if (-not $isAdmin) {
        Log 'could not stop the old server without elevation - relaunching elevated (UAC prompt)'
        Start-Process powershell.exe -Verb RunAs -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-NoExit', '-File', ('"' + $me + '"'))
        exit 0
    }
    Log 'could not stop the old server even elevated - inspect it by hand'
    exit 1
}

$env:PORT = "$port"
Start-Process -FilePath 'cmd.exe' -ArgumentList @('/k', 'start_server.bat') -WorkingDirectory $repo
$new = @()
for ($i = 0; $i -lt 60; $i++) {
    Start-Sleep -Seconds 1
    $new = @(Port-Pids $port)
    if ($new.Count -gt 0) { break }
}
if ($new.Count -eq 0) { Log ('no listener on :' + $port + ' after 60 s - look at the launcher window'); exit 1 }
$r = Invoke-WebRequest -Uri ('http://127.0.0.1:' + $port + '/login') -UseBasicParsing -TimeoutSec 30
Log (':' + $port + ' is up - pid ' + ($new -join ',') + ' - HTTP ' + $r.StatusCode)
