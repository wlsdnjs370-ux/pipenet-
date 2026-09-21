#Requires -RunAsAdministrator
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$statusPath = Join-Path $PSScriptRoot '_deploy_module_f_cancel_5051.json'
$expectedOldPid = 45560
function Port-Pids([int]$number) {
    @(netstat -ano -p tcp | ForEach-Object {
        if ($_ -match "^\s*TCP\s+\S+:$number\s+\S+\s+LISTENING\s+(\d+)\s*$") {
            [int]$Matches[1]
        }
    } | Sort-Object -Unique)
}
try {
    $domainBefore = @(Port-Pids 5052)
    $localBefore = @(Port-Pids 5051)
    if ($localBefore.Count -ne 1 -or $localBefore[0] -ne $expectedOldPid) {
        throw 'Port 5051 process changed; inspect it before restarting. No process was stopped.'
    }
    $old = Get-Process -Id $expectedOldPid
    if ($old.ProcessName -ne 'python') { throw 'Expected Python on 5051. No process was stopped.' }
    Stop-Process -Id $expectedOldPid -Force
    $localAfter = @()
    for ($attempt=0; $attempt -lt 15; $attempt++) {
        Start-Sleep -Seconds 1
        $localAfter = @(Port-Pids 5051)
        if ($localAfter.Count -gt 0) { break }
    }
    if ($localAfter.Count -eq 0) {
        $env:PORT = '5051'
        $env:HOST = '0.0.0.0'
        $env:PYTHONUTF8 = '1'
        $python = 'C:\Users\admin\anaconda3\python.exe'
        Start-Process -FilePath $python -ArgumentList @('-u','serve.py') -WorkingDirectory $repo -WindowStyle Hidden `
            -RedirectStandardOutput (Join-Path $PSScriptRoot '_local_5051_cancel.stdout.log') `
            -RedirectStandardError (Join-Path $PSScriptRoot '_local_5051_cancel.stderr.log') | Out-Null
        for ($attempt=0; $attempt -lt 25; $attempt++) {
            Start-Sleep -Seconds 1
            $localAfter = @(Port-Pids 5051)
            if ($localAfter.Count -gt 0) { break }
        }
    }
    if ($localAfter.Count -ne 1 -or $localAfter[0] -eq $expectedOldPid) { throw 'No unique new listener on 5051.' }
    $domainAfter = @(Port-Pids 5052)
    if (($domainBefore -join ',') -ne ($domainAfter -join ',')) { throw '5052 process changed externally; inspect server status.' }
    $response = Invoke-WebRequest -Uri 'http://127.0.0.1:5051/login' -UseBasicParsing -TimeoutSec 15
    if ($response.StatusCode -ne 200) { throw 'New 5051 server did not respond successfully.' }
    @{Ok=$true;OldPid=$expectedOldPid;NewPid=$localAfter[0];DomainPid=$domainAfter;Time=(Get-Date).ToString('o')} |
        ConvertTo-Json | Set-Content -LiteralPath $statusPath -Encoding UTF8
} catch {
    @{Ok=$false;Error=$_.Exception.Message;Time=(Get-Date).ToString('o')} |
        ConvertTo-Json | Set-Content -LiteralPath $statusPath -Encoding UTF8
    exit 1
}


