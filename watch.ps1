# Auto-push ke Google Apps Script setiap ada perubahan file
# Jalankan: klik kanan file ini -> Run with PowerShell

$projectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectDir

Write-Host "===========================================" -ForegroundColor Cyan
Write-Host "  GAS Auto-Push Watcher - gas-baru" -ForegroundColor Cyan
Write-Host "  Setiap file berubah -> langsung push ke GAS" -ForegroundColor Cyan
Write-Host "  Tekan Ctrl+C untuk berhenti" -ForegroundColor Cyan
Write-Host "===========================================" -ForegroundColor Cyan
Write-Host ""

clasp push --watch
