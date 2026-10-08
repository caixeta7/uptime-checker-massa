function Test-ComputerUptime {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName
    )

    # Online/Offline rápido via Test-Connection
    if (-not (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
        return [PSCustomObject]@{
            Computador = $ComputerName
            Status     = 'OFFLINE'
            Usuario    = ''
            BootTime   = ''
            Uptime     = 'N/A'
            DiasLigado = ''
            HorasLigado= ''
        }
    }

    try {
        $OS = Get-WmiObject Win32_OperatingSystem -ComputerName $ComputerName -ErrorAction Stop

        $CS = Get-WmiObject Win32_ComputerSystem -ComputerName $ComputerName -ErrorAction SilentlyContinue
        $LoggedUser = $CS.UserName
        if ([string]::IsNullOrWhiteSpace($LoggedUser)) {
            $LoggedUser = 'SEM USUARIO'
        }

        $Boot = $OS.ConvertToDateTime($OS.LastBootUpTime)
        $Uptime = (Get-Date) - $Boot

        return [PSCustomObject]@{
            Computador = $ComputerName
            Status     = 'ONLINE'
            Usuario    = $LoggedUser
            BootTime   = $Boot.ToString('dd/MM/yyyy HH:mm:ss')
            Uptime     = "$([int]$Uptime.TotalDays)d $([int]($Uptime.TotalHours % 24))h $([int]($Uptime.TotalMinutes % 60))m"
            DiasLigado = [int]$Uptime.TotalDays
            HorasLigado= [int]$Uptime.TotalHours
        }
    }
    catch {
        return [PSCustomObject]@{
            Computador = $ComputerName
            Status     = 'ERRO WMI'
            Usuario    = ''
            BootTime   = ''
            Uptime     = 'N/A'
            DiasLigado = ''
            HorasLigado= ''
        }
    }
}
