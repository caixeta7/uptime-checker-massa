<#
================================================================================
ORQUESTRADOR SYSTEMINFO - NATIVO MODULAR
================================================================================
#>

$ProgressPreference = 'SilentlyContinue'
$PSScriptRootPath = $PSScriptRoot

# Importa Configurações
$ConfigFile = "$PSScriptRootPath\Config\settings.psd1"
if (Test-Path $ConfigFile) {
    $Config = Import-PowerShellDataFile -Path $ConfigFile
} else {
    $Config = @{ MaxThreads = 80 }
}

# Importa Módulos
Import-Module "$PSScriptRootPath\Modules\UI.Console.psm1" -Force
Import-Module "$PSScriptRootPath\Modules\Core.Network.psm1" -Force
Import-Module "$PSScriptRootPath\Modules\Core.Excel.psm1" -Force

Write-Header "SYSTEMINFO ULTRA - VERIFICACAO DE UPTIME"

$PcsFile = "$PSScriptRootPath\pcs.txt"
$RelatoriosPath = "$PSScriptRootPath\Relatorios"

if (-not (Test-Path $RelatoriosPath)) {
    New-Item -ItemType Directory -Path $RelatoriosPath | Out-Null
}

if (-not (Test-Path $PcsFile)) {
    Write-StatusMessage "Error" "Arquivo pcs.txt nao encontrado: $PcsFile"
    exit 1
}

$Computers = Get-Content $PcsFile | Where-Object { $_.Trim() -ne "" } | ForEach-Object { $_.Trim().ToUpper() } | Select-Object -Unique

if (-not $Computers -or $Computers.Count -eq 0) {
    Write-StatusMessage "Warning" "Nenhuma maquina listada no pcs.txt"
    exit 1
}

Write-StatusMessage "Info" "Carregados $($Computers.Count) computadores. Iniciando varredura paralela..."

# Runspace Pool para máxima performance
$Pool = [RunspaceFactory]::CreateRunspacePool(1, [int]$Config.MaxThreads)
$Pool.Open()

$Jobs = foreach ($Computer in $Computers) {
    $ScriptBlock = {
        param($Comp, $Path)
        Import-Module "$Path\Modules\Core.Network.psm1" -Force
        Test-ComputerUptime -ComputerName $Comp
    }

    $PowerShell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($Computer).AddArgument($PSScriptRootPath)
    $PowerShell.RunspacePool = $Pool
    [PSCustomObject]@{
        Pipe   = $PowerShell
        Handle = $PowerShell.BeginInvoke()
    }
}

# Coleta de Resultados
$Results = foreach ($Job in $Jobs) {
    try {
        $Job.Pipe.EndInvoke($Job.Handle)
    }
    catch { }
    $Job.Pipe.Dispose()
}

$Pool.Close()
$Pool.Dispose()

if (-not $Results) {
    Write-StatusMessage "Warning" "Nenhum resultado retornado."
    exit 1
}

# Saída visual no console
foreach ($res in $Results) {
    if ($res.Status -eq 'ONLINE') {
        Write-Host " [ONLINE]  $($res.Computador) - Uptime: $($res.Uptime) - User: $($res.Usuario)" -ForegroundColor Green
    }
    elseif ($res.Status -eq 'ERRO WMI') {
        Write-Host " [ERRO WMI] $($res.Computador)" -ForegroundColor Yellow
    }
    else {
        Write-Host " [OFFLINE] $($res.Computador)" -ForegroundColor DarkGray
    }
}

# Exportação
$DateStamp = Get-Date -Format "yyyy-MM-dd_HH-mm"
$OutputExcel = "$RelatoriosPath\Uptime_Maquinas_$DateStamp.xlsx"

Write-StatusMessage "Info" "Gerando relatorio Excel..."
$finalPath = Write-ExcelReport -Data $Results -OutputPath $OutputExcel

Write-Host ""
Write-StatusMessage "Success" "PROCESSO CONCLUIDO COM SUCESSO"
Write-StatusMessage "Info" "Relatorio salvo em: $finalPath"
Write-StatusMessage "Info" "Maquinas analisadas: $($Computers.Count)"
Write-StatusMessage "Info" "Resultados retornados: $($Results.Count)"
Write-Host ""
