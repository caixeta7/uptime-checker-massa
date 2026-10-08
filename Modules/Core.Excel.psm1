function Write-ExcelReport {
    param (
        [Parameter(Mandatory = $true)]
        [array]$Data,
        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    if (-not (Get-Module -ListAvailable ImportExcel)) {
        Write-Warning "Módulo ImportExcel não encontrado. Exportando em CSV simples."
        $csvPath = $OutputPath -replace '\.xlsx$', '.csv'
        $Data | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        return $csvPath
    }

    $Data | Export-Excel -Path $OutputPath -WorksheetName "Uptime" -AutoSize -AutoFilter -BoldTopRow -FreezePane 2,1 -TableName "UptimeTable" -TableStyle Medium2 -ClearSheet

    # Aplicação de Coloração/Destaque por Uptime
    $Excel = Open-ExcelPackage $OutputPath
    $Sheet = $Excel.Workbook.Worksheets["Uptime"]
    $LastRow = $Sheet.Dimension.End.Row

    for ($Row = 2; $Row -le $LastRow; $Row++) {
        $Dias = $Sheet.Cells["F$Row"].Value # Coluna DiasLigado

        if ($Dias -ne $null -and $Dias -ge 2) {
            $Sheet.Cells["A$Row:G$Row"].Style.Fill.PatternType = "Solid"
            $Sheet.Cells["A$Row:G$Row"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightCoral)
        }
        elseif ($Dias -ne $null -and $Dias -ge 1) {
            $Sheet.Cells["A$Row:G$Row"].Style.Fill.PatternType = "Solid"
            $Sheet.Cells["A$Row:G$Row"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Khaki)
        }
    }

    Close-ExcelPackage $Excel
    return $OutputPath
}
