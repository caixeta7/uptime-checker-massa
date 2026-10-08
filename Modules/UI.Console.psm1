function Write-Header {
    param ([string]$Title)
    Clear-Host
    Write-Host "========================================================" -ForegroundColor Cyan
    Write-Host " $Title" -ForegroundColor Yellow
    Write-Host "========================================================" -ForegroundColor Cyan
    Write-Host ""
}

function Write-StatusMessage {
    param (
        [string]$Type,
        [string]$Message
    )
    switch ($Type) {
        "Info"    { Write-Host "[INFO] $Message" -ForegroundColor Cyan }
        "Success" { Write-Host "[SUCESSO] $Message" -ForegroundColor Green }
        "Warning" { Write-Host "[AVISO] $Message" -ForegroundColor Yellow }
        "Error"   { Write-Host "[ERRO] $Message" -ForegroundColor Red }
    }
}
