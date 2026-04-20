# Script que mantiene CACHANPESCA_CapturarEtiquetas siempre running
$exePath = "$PSScriptRoot\CACHANPESCA_CapturarEtiquetas.exe"
$lockFile = "$PSScriptRoot\keep_alive.lock"
$procName = "CACHANPESCA_CapturarEtiquetas"

if (Test-Path $lockFile) {
    $lockPid = Get-Content $lockFile -ErrorAction SilentlyContinue
    if ($lockPid -and (Get-Process -Id $lockPid -ErrorAction SilentlyContinue)) {
        exit
    }
    Remove-Item $lockFile -Force -ErrorAction SilentlyContinue
}
$PID | Out-File $lockFile -Encoding ASCII

while ($true) {
    $procs = Get-Process -Name $procName -ErrorAction SilentlyContinue
    if (-not $procs) {
        Start-Process $exePath -WindowStyle Hidden
    } elseif ($procs.Count -gt 1) {
        $procs | Select-Object -Skip 1 | Stop-Process -Force -ErrorAction SilentlyContinue
    }
}
