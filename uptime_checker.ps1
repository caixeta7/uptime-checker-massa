<#
================================================================================
UPTIME CHECKER ULTRA RÁPIDO
================================================================================

FOCO:
- Máximo desempenho
- Baixo consumo
- Ignora offline imediatamente
- Consulta somente o necessário
- Exporta Excel organizado
- Nome do arquivo com data/hora
- Destaque visual por uptime

REQUISITOS:
Install-Module ImportExcel -Scope CurrentUser -Force

ARQUIVOS:
- pcs.txt
- uptime.ps1

================================================================================
#>

# =========================
# PERFORMANCE
# =========================

$ProgressPreference = 'SilentlyContinue'

# =========================
# CONFIG
# =========================

$PcsFile = "$PSScriptRoot\pcs.txt"

# Pasta de relatórios
$RelatoriosPath = "$PSScriptRoot\Relatorios"

if (-not (Test-Path $RelatoriosPath)) {
    New-Item -ItemType Directory -Path $RelatoriosPath | Out-Null
}

# Data/hora no nome do arquivo
$DateStamp = Get-Date -Format "yyyy-MM-dd_HH-mm"

$OutputExcel = "$RelatoriosPath\Uptime_Maquinas_$DateStamp.xlsx"

# Threads
$MaxThreads = 80

# =========================
# VALIDA TXT
# =========================

if (-not (Test-Path $PcsFile)) {

    Write-Host ''
    Write-Host 'Arquivo pcs.txt não encontrado.'
    Write-Host $PcsFile
    Write-Host ''

    exit
}

# =========================
# LEITURA DAS MÁQUINAS
# =========================

$Computers = Get-Content $PcsFile |
Where-Object {
    $_.Trim() -ne ""
} |
ForEach-Object {
    $_.Trim().ToUpper()
} |
Sort-Object -Unique

if ($Computers.Count -eq 0) {

    Write-Host ''
    Write-Host 'Nenhuma máquina encontrada no pcs.txt'
    Write-Host ''

    exit
}

# =========================
# RUNSPACE POOL
# =========================

$Pool = [RunspaceFactory]::CreateRunspacePool(1, $MaxThreads)
$Pool.Open()

$Jobs = foreach ($Computer in $Computers) {

    $PowerShell = [powershell]::Create()

    $PowerShell.RunspacePool = $Pool

    [void]$PowerShell.AddScript({

        param($Computer)

        # =====================================================
        # TESTE RÁPIDO ONLINE/OFFLINE
        # =====================================================

        if (-not (Test-Connection `
            -ComputerName $Computer `
            -Count 1 `
            -Quiet `
            -ErrorAction SilentlyContinue)) {

            return [PSCustomObject]@{
                Hostname    = $Computer
                Status      = "OFFLINE"
                Usuario     = ""
                BootTime    = ""
                DiasLigado  = ""
                HorasLigado = ""
            }
        }

        try {

            # =================================================
            # SISTEMA OPERACIONAL
            # =================================================

            $OS = Get-WmiObject `
                Win32_OperatingSystem `
                -ComputerName $Computer `
                -ErrorAction Stop

            # =================================================
            # USUÁRIO LOGADO
            # =================================================

            $CS = Get-WmiObject `
                Win32_ComputerSystem `
                -ComputerName $Computer `
                -ErrorAction SilentlyContinue

            $LoggedUser = $CS.UserName

            if ([string]::IsNullOrWhiteSpace($LoggedUser)) {
                $LoggedUser = "SEM USUARIO"
            }

            # =================================================
            # UPTIME
            # =================================================

            $Boot = $OS.ConvertToDateTime($OS.LastBootUpTime)

            $Uptime = (Get-Date) - $Boot

            return [PSCustomObject]@{
                Hostname    = $Computer
                Status      = "ONLINE"
                Usuario     = $LoggedUser
                BootTime    = $Boot.ToString("dd/MM/yyyy HH:mm:ss")
                DiasLigado  = [int]$Uptime.TotalDays
                HorasLigado = [int]$Uptime.TotalHours
            }
        }
        catch {

            return [PSCustomObject]@{
                Hostname    = $Computer
                Status      = "ERRO WMI"
                Usuario     = ""
                BootTime    = ""
                DiasLigado  = ""
                HorasLigado = ""
            }
        }

    }).AddArgument($Computer)

    [PSCustomObject]@{
        Pipe   = $PowerShell
        Handle = $PowerShell.BeginInvoke()
    }
}

# =========================
# COLETA RESULTADOS
# =========================

$Results = foreach ($Job in $Jobs) {

    try {
        $Job.Pipe.EndInvoke($Job.Handle)
    }
    catch {
    }

    $Job.Pipe.Dispose()
}

# =========================
# FINALIZA RUNSPACES
# =========================

$Pool.Close()
$Pool.Dispose()

# =========================
# ORDENA RESULTADOS
# =========================

$Results = $Results |
Sort-Object Hostname

# =========================
# GARANTE RESULTADO
# =========================

if (-not $Results) {

    Write-Host ''
    Write-Host 'Nenhum resultado retornado.'
    Write-Host ''

    exit
}

# =========================
# IMPORTA MÓDULO EXCEL
# =========================

try {

    Import-Module ImportExcel -ErrorAction Stop
}
catch {

    Write-Host ''
    Write-Host 'Módulo ImportExcel não encontrado.'
    Write-Host ''
    Write-Host 'Execute:'
    Write-Host 'Install-Module ImportExcel -Scope CurrentUser -Force'
    Write-Host ''

    exit
}

# =========================
# EXPORTA EXCEL
# =========================

$Results |
Export-Excel `
    -Path $OutputExcel `
    -WorksheetName "Uptime" `
    -AutoSize `
    -FreezeTopRow `
    -BoldTopRow `
    -TableName "UptimeTable" `
    -TableStyle Medium2

# =========================
# ABRE PLANILHA
# =========================

$Excel = Open-ExcelPackage $OutputExcel
$Sheet = $Excel.Workbook.Worksheets["Uptime"]

# =========================
# COLORAÇÃO
# =========================

$LastRow = $Sheet.Dimension.End.Row

for ($Row = 2; $Row -le $LastRow; $Row++) {

    $Dias = $Sheet.Cells["E$Row"].Value

    # >= 2 dias = vermelho
    if ($Dias -ge 2) {

        $Sheet.Cells["A$Row:F$Row"].Style.Fill.PatternType = "Solid"

        $Sheet.Cells["A$Row:F$Row"].Style.Fill.BackgroundColor.SetColor(
            [System.Drawing.Color]::LightCoral
        )
    }

    # >= 1 dia = amarelo
    elseif ($Dias -ge 1) {

        $Sheet.Cells["A$Row:F$Row"].Style.Fill.PatternType = "Solid"

        $Sheet.Cells["A$Row:F$Row"].Style.Fill.BackgroundColor.SetColor(
            [System.Drawing.Color]::Khaki
        )
    }
}

# =========================
# SALVA EXCEL
# =========================

Close-ExcelPackage $Excel

# =========================
# FINAL
# =========================

Write-Host ''
Write-Host '====================================='
Write-Host 'RELATÓRIO GERADO COM SUCESSO'
Write-Host '====================================='
Write-Host ''
Write-Host $OutputExcel
Write-Host ''
Write-Host ('Máquinas analisadas: ' + $Computers.Count)
Write-Host ('Resultados retornados: ' + $Results.Count)
Write-Host ''