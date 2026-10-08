# Uptime Checker em Massa

Script PowerShell para verificar uptime, usuário logado e informações de sistema de múltiplas máquinas Windows simultaneamente, exportando relatório Excel formatado com destaque visual por tempo de atividade.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue) ![Windows](https://img.shields.io/badge/Windows-WMI-lightgrey)

## Funcionalidades

- **Paralelismo via RunspacePool** com até 80 threads simultâneas — processa centenas de máquinas em segundos
- **Ignora offline imediatamente**: ping rápido antes de tentar WMI, sem timeout desnecessário
- **Coleta via WMI**: uptime, último boot, usuário logado, hostname real, OS
- **Exportação Excel** com módulo `ImportExcel`: nome do arquivo inclui data/hora automaticamente
- **Formatação condicional**: destaque visual por faixa de uptime (horas, dias, semanas)
- **Pasta de relatórios** criada automaticamente em `.\Relatorios\`
- **Lista de hosts via arquivo externo** (`pcs.txt`): sem edição de código

## Como usar

### 1. Instalar dependência

```powershell
Install-Module ImportExcel -Scope CurrentUser -Force
```

### 2. Configurar lista de máquinas

```powershell
Copy-Item pcs.txt.example pcs.txt
# Edite pcs.txt com os hostnames do seu ambiente (um por linha)
```

### 3. Executar

```powershell
.\Run-Uptime.ps1
```

Ou dê dois cliques em `uptime.bat`. Threads, timeout e quantidade de pings ficam em `Config\settings.psd1`.

O relatório é gerado automaticamente em `.\Relatorios\Uptime_Maquinas_YYYY-MM-DD_HH-mm.xlsx` (ou CSV, se o `ImportExcel` não estiver instalado).

## Formato do pcs.txt

```
CORP-SP-001
CORP-SP-002
CORP-RJ-001
```

## Estrutura

```
Run-Uptime.ps1          orquestrador
uptime.bat              atalho para executar
Config/settings.psd1    MaxThreads, TimeoutMs, PingCount
Modules/
  Core.Network.psm1     ping + coleta WMI em RunspacePool
  Core.Excel.psm1       exportação Excel (fallback CSV)
  UI.Console.psm1       saída formatada no console
```

## Pré-requisitos

- Windows com PowerShell 5.1+
- Acesso WMI nas máquinas-alvo (admin de domínio ou conta de serviço)
- Módulo `ImportExcel` (opcional; sem ele o relatório sai em CSV)

## Contexto

Desenvolvido para auditoria periódica de parque de máquinas corporativo — geração rápida de relatório executivo de disponibilidade e presença de usuários, útil para equipes de suporte e gestão de ativos.
