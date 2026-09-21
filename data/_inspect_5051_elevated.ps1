$ErrorActionPreference = 'Stop'
$resultPath = Join-Path $PSScriptRoot '_inspect_5051_elevated.json'
try {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    $taskRows = @(Get-ScheduledTask | Where-Object {
        $_.TaskName -match 'FNCAD|5051|5052' -or
        ($_.Actions.Arguments -join ' ') -match 'serve.py'
    } | ForEach-Object {
        [pscustomobject]@{
            Name=$_.TaskName; Path=$_.TaskPath; State=[string]$_.State
            Actions=@($_.Actions | Select-Object Execute,Arguments,WorkingDirectory)
        }
    })
    $procs = @(Get-CimInstance Win32_Process -Filter 'ProcessId=4080 OR ProcessId=4048' |
        Select-Object ProcessId,ParentProcessId,Name,ExecutablePath,CommandLine)
    [pscustomobject]@{
        Elevated=$principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        Tasks=$taskRows; Processes=$procs
    } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $resultPath -Encoding UTF8
} catch {
    @{Error=$_.Exception.Message} | ConvertTo-Json | Set-Content -LiteralPath $resultPath -Encoding UTF8
    exit 1
}
