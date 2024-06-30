# Definir politica de execucao temporariamente
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process -Force

# Definir codificacao do console para UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Criar pasta para logs
$logPath = "C:\MS365"
if (-Not (Test-Path -Path $logPath)) {
    New-Item -Path $logPath -ItemType Directory
}

# Caminho do arquivo de log
$logFilePath = "$logPath\error_log.txt"

# Funcao para logar erros
function Log-Error {
    param (
        [string]$message
    )
    Add-Content -Path $logFilePath -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $message"
}

# Funcao para verificar e instalar/atualizar modulo Microsoft Graph
function Check-Install-Modules {
    $moduleName = "Microsoft.Graph"
    try {
        if (-not (Get-Module -Name $moduleName -ListAvailable)) {
            Write-Host "Modulo $moduleName nao encontrado. Tentando instalar..." -ForegroundColor Yellow
            Install-Module -Name $moduleName -Scope CurrentUser -Force -ErrorAction Stop
        } else {
            Write-Host "Modulo $moduleName ja esta instalado." -ForegroundColor Green
        }
    } catch {
        $errorMessage = "Erro ao instalar/atualizar o modulo ${moduleName}: $($_.Exception.Message)"
        Write-Host $errorMessage -ForegroundColor Red
        Log-Error $errorMessage
    }
}

# Tela de boas-vindas
function Show-WelcomeScreen {
    Clear-Host
    Write-Host "===========================================" -ForegroundColor Green
    Write-Host "              JORNADA365                  " -ForegroundColor Green
    Write-Host "            Sua Jornada Comenca Aqui       " -ForegroundColor Green
    Write-Host "===========================================" -ForegroundColor Green
    Write-Host "Este script foi criado para simplificar o gerenciamento de licencas no Microsoft 365." -ForegroundColor Yellow
    Write-Host "Gerenciar licencas e uma tarefa facil quando se trata de poucas licencas ou mesmo" -ForegroundColor Yellow
    Write-Host "algumas dezenas de usuarios. No entanto, ao remover ou substituir licencas para" -ForegroundColor Yellow
    Write-Host "centenas de usuarios, o processo se torna muito mais complexo." -ForegroundColor Yellow
    Write-Host "Por isso, utilize este script com cautela e sinta-se a vontade para compartilha-lo." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Visite nosso site: www.jornada365.cloud" -ForegroundColor Cyan
    Write-Host "Microsoft 365: admin.microsoft.com" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "SEJA BEM-VINDO || JORNADA 365" -ForegroundColor Green
    Write-Host "===========================================" -ForegroundColor Green
}

# Conectar ao Microsoft Graph com mecanismo de repeticao
function Connect-MicrosoftGraph {
    $maxRetries = 5
    $retryCount = 0
    $connected = $false

    while (-not $connected -and $retryCount -lt $maxRetries) {
        try {
            Write-Host "Conectando ao Microsoft Graph..." -ForegroundColor Cyan
            Connect-MgGraph -Scopes "Directory.ReadWrite.All", "User.ReadWrite.All", "Group.ReadWrite.All", "Sites.ReadWrite.All", "DeviceManagementManagedDevices.ReadWrite.All", "Reports.Read.All", "Mail.ReadWrite"
            Write-Host "Conectado ao Microsoft Graph." -ForegroundColor Green
            $connected = $true
        } catch {
            $retryCount++
            $errorMessage = "Erro de autenticacao: $($_.Exception.Message). Tentativa $retryCount de $maxRetries."
            Write-Host $errorMessage -ForegroundColor Red
            Log-Error $errorMessage
            Start-Sleep -Seconds 10
        }
    }

    if (-not $connected) {
        Write-Host "Falha ao conectar ao Microsoft Graph apos $maxRetries tentativas." -ForegroundColor Red
        exit
    }
}

# Desconectar do Microsoft Graph
function Disconnect-Services {
    Write-Host "Desconectando do Microsoft Graph..." -ForegroundColor Cyan
    try {
        Disconnect-MgGraph -ErrorAction Stop
        Write-Host "Desconectado do Microsoft Graph." -ForegroundColor Green
    } catch {
        Write-Host "Nenhuma aplicacao para desconectar." -ForegroundColor Yellow
    }
}

# Funcao para obter nomes amigaveis de SKU
function Get-FriendlySkuNames {
    $skuUrl = "https://raw.githubusercontent.com/MicrosoftDocs/entra-docs/main/docs/identity/users/licensing-service-plan-reference.md"
    $skuNames = @{}

    try {
        $content = Invoke-WebRequest -Uri $skuUrl -UseBasicParsing
        $lines = $content.Content -split "`n"
        foreach ($line in $lines) {
            if ($line -match "\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|") {
                $skuNames[$matches[3]] = $matches[1] # Usando GUID como chave e nome do produto como valor
            }
        }
    } catch {
        $errorMessage = "Erro ao obter nomes amigaveis de SKU: $($_.Exception.Message)"
        Write-Host $errorMessage -ForegroundColor Red
        Log-Error $errorMessage
    }
    return $skuNames
}

# Listar licencas disponiveis
function Get-AvailableSkus {
    [array]$Skus = Get-MgSubscribedSku
    $SkuList = [System.Collections.Generic.List[Object]]::new()
    $friendlySkuNames = Get-FriendlySkuNames

    foreach ($Sku in $Skus) {
        $SkuAvailable = ($Sku.PrepaidUnits.Enabled - $Sku.ConsumedUnits)
        $SkuName = $friendlySkuNames[$Sku.SkuId] # Procurar nome amigavel usando GUID
        $ReportLine = [PSCustomObject]@{
            SkuId         = $Sku.SkuId
            SkuPartNumber = if ($SkuName) { $SkuName } else { $Sku.SkuPartNumber }
            Consumido     = $Sku.ConsumedUnits
            Pago          = $Sku.PrepaidUnits.Enabled
            Disponivel    = $SkuAvailable
        }
        $SkuList.Add($ReportLine)
    }

    # Remover SKUs sem licencas disponiveis
    $SkuList = $SkuList | Where-Object {$_.Disponivel -gt 0}
    return $SkuList
}

# Importar usuarios de um arquivo CSV
function Import-UsersFromCsv {
    Add-Type -AssemblyName System.Windows.Forms
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Filter = "Arquivos CSV (*.csv)|*.csv|Todos os arquivos (*.*)|*.*"
    $fileDialog.ShowDialog() | Out-Null
    $filePath = $fileDialog.FileName

    if ($filePath) {
        return Import-Csv -Path $filePath
    } else {
        Write-Host "Nenhum arquivo selecionado." -ForegroundColor Red
        return $null
    }
}

# Atribuir licencas
function Assign-Licenses {
    param (
        [array]$skuIds,
        [array]$users
    )
    $AssignmentReport = [System.Collections.Generic.List[PSCustomObject]]::new()
    $i = 0
    $friendlySkuNames = Get-FriendlySkuNames
    foreach ($User in $Users) {
        $ErrorMsg = $Null; $i++
        Write-Host ("Processando conta de $i/$($Users.Count)") -ForegroundColor Cyan
        try {
            $UserData = Get-MgUser -UserId $User.Email.Trim() -Property id, assignedLicenses, department, displayName -ErrorAction Stop
            $DisplayName = $UserData.DisplayName
        } catch {
            $ErrorMsg = "Erro ao buscar dados do usuario ${User.Email}: $($_.Exception.Message)"
            Write-Host $ErrorMsg -ForegroundColor Red
            $AssignmentReport.Add([PSCustomObject]@{
                Numero                      = $i
                Nome                        = $User.DisplayName
                Email                       = $User.Mail
                Departamento                = $User.Department
                Licenca                     = $friendlySkuNames[$skuId]
                "Data/Hora da execucao"     = (Get-Date -format "dd/MM/yyyy HH:mm:ss")
                Erro                        = $ErrorMsg
            })
            Log-Error $ErrorMsg
            continue
        }

        foreach ($skuId in $skuIds) {
            $LicenseData = $UserData | Select-Object -ExpandProperty AssignedLicenses
            if ($skuId -in $LicenseData.SkuId) {
                $StatusMsg = "Licenca ja atribuida a conta de usuario ${User.Email}"
                Write-Host $StatusMsg -ForegroundColor Yellow
                $AssignmentReport.Add([PSCustomObject]@{
                    Numero                      = $i
                    Nome                        = $DisplayName
                    Email                       = $User.Mail
                    Departamento                = $User.Department
                    Licenca                     = $friendlySkuNames[$skuId]
                    "Data/Hora da execucao"     = (Get-Date -format "dd/MM/yyyy HH:mm:ss")
                    Erro                        = $StatusMsg
                })
            } else {
                try {
                    Set-MgUserLicense -UserId $User.Email -Addlicenses @{SkuId = $skuId} -RemoveLicenses @() -ErrorAction Stop
                    $StatusMsg = "Licenca atribuida $($friendlySkuNames[$skuId]) - $($DisplayName)"
                    Write-Host "OK - " -ForegroundColor Green -NoNewline
                    Write-Host $StatusMsg -ForegroundColor Green
                    $AssignmentReport.Add([PSCustomObject]@{
                        Numero                      = $i
                        Nome                        = $DisplayName
                        Email                       = $User.Mail
                        Departamento                = $User.Department
                        Licenca                     = $friendlySkuNames[$skuId]
                        "Data/Hora da execucao"     = (Get-Date -format "dd/MM/yyyy HH:mm:ss")
                    })
                } catch {
                    $ErrorMsg = "Erro ao atribuir licenca ao usuario ${User.Email}: $($_.Exception.Message)"
                    Write-Host $ErrorMsg -ForegroundColor Red
                    $AssignmentReport.Add([PSCustomObject]@{
                        Numero                      = $i
                        Nome                        = $DisplayName
                        Email                       = $User.Mail
                        Departamento                = $User.Department
                        Licenca                     = $friendlySkuNames[$skuId]
                        "Data/Hora da execucao"     = (Get-Date -format "dd/MM/yyyy HH:mm:ss")
                        Erro                        = $ErrorMsg
                    })
                    Log-Error $ErrorMsg
                }
            }
        }
    }
    return $AssignmentReport
}

# Funcao para remover licencas
function Remove-Licenses {
    param (
        [array]$skuIds,
        [array]$users
    )
    $RemovalReport = [System.Collections.Generic.List[PSCustomObject]]::new()
    $i = 0
    $friendlySkuNames = Get-FriendlySkuNames
    foreach ($User in $Users) {
        $ErrorMsg = $Null; $i++
        Write-Host ("Processando conta de $i/$($Users.Count)") -ForegroundColor Cyan
        try {
            $UserData = Get-MgUser -UserId $User.Email.Trim() -Property id, assignedLicenses, department, displayName -ErrorAction Stop
            $DisplayName = $UserData.DisplayName
        } catch {
            $ErrorMsg = "Erro ao buscar dados do usuario ${User.Email}: $($_.Exception.Message)"
            Write-Host $ErrorMsg -ForegroundColor Red
            $RemovalReport.Add([PSCustomObject]@{
                Numero                      = $i
                Nome                        = $User.DisplayName
                Email                       = $User.Mail
                Departamento                = $User.Department
                Licenca                     = $friendlySkuNames[$skuId]
                "Data/Hora da execucao"     = (Get-Date -format "dd/MM/yyyy HH:mm:ss")
                Erro                        = $ErrorMsg
            })
            Log-Error $ErrorMsg
            continue
        }

        foreach ($skuId in $skuIds) {
            $LicenseData = $UserData | Select-Object -ExpandProperty AssignedLicenses
            if ($skuId -in $LicenseData.SkuId) {
                try {
                    Set-MgUserLicense -UserId $User.Email -Addlicenses @() -RemoveLicenses @($skuId) -ErrorAction Stop
                    $StatusMsg = "Licenca removida $($friendlySkuNames[$skuId]) - $($DisplayName)"
                    Write-Host "OK - " -ForegroundColor Green -NoNewline
                    Write-Host $StatusMsg -ForegroundColor Green
                    $RemovalReport.Add([PSCustomObject]@{
                        Numero                      = $i
                        Nome                        = $DisplayName
                        Email                       = $User.Mail
                        Departamento                = $User.Department
                        Licenca                     = $friendlySkuNames[$skuId]
                        "Data/Hora da execucao"     = (Get-Date -format "dd/MM/yyyy HH:mm:ss")
                    })
                } catch {
                    $ErrorMsg = "Erro ao remover licenca do usuario ${User.Email}: $($_.Exception.Message)"
                    Write-Host $ErrorMsg -ForegroundColor Red
                    $RemovalReport.Add([PSCustomObject]@{
                        Numero                      = $i
                        Nome                        = $DisplayName
                        Email                       = $User.Mail
                        Departamento                = $User.Department
                        Licenca                     = $friendlySkuNames[$skuId]
                        "Data/Hora da execucao"     = (Get-Date -format "dd/MM/yyyy HH:mm:ss")
                        Erro                        = $ErrorMsg
                    })
                    Log-Error $ErrorMsg
                }
            } else {
                $StatusMsg = "Licenca nao atribuida a conta de usuario ${User.Email}"
                Write-Host $StatusMsg -ForegroundColor Yellow
                $RemovalReport.Add([PSCustomObject]@{
                    Numero                      = $i
                    Nome                        = $DisplayName
                    Email                       = $User.Mail
                    Departamento                = $User.Department
                    Licenca                     = $friendlySkuNames[$skuId]
                    "Data/Hora da execucao"     = (Get-Date -format "dd/MM/yyyy HH:mm:ss")
                    Erro                        = $StatusMsg
                })
            }
        }
    }
    return $RemovalReport
}

# Funcao para remover todas as licencas
function Remove-AllLicenses {
    param (
        [array]$users
    )
    $RemovalReport = [System.Collections.Generic.List[PSCustomObject]]::new()
    $i = 0
    foreach ($User in $Users) {
        $ErrorMsg = $Null; $i++
        Write-Host ("Processando conta de $i/$($Users.Count)") -ForegroundColor Cyan
        try {
            $UserData = Get-MgUser -UserId $User.Email.Trim() -Property id, assignedLicenses, department, displayName -ErrorAction Stop
            $DisplayName = $UserData.DisplayName
        } catch {
            $ErrorMsg = "Erro ao buscar dados do usuario ${User.Email}: $($_.Exception.Message)"
            Write-Host $ErrorMsg -ForegroundColor Red
            $RemovalReport.Add([PSCustomObject]@{
                Numero                      = $i
                Nome                        = $User.DisplayName
                Email                       = $User.Mail
                Departamento                = $User.Department
                Licenca                     = "Todas"
                "Data/Hora da execucao"     = (Get-Date -format "dd/MM/yyyy HH:mm:ss")
                Erro                        = $ErrorMsg
            })
            Log-Error $ErrorMsg
            continue
        }

        $LicenseData = $UserData | Select-Object -ExpandProperty AssignedLicenses
        if ($LicenseData.Count -gt 0) {
            try {
                Set-MgUserLicense -UserId $User.Email -Addlicenses @() -RemoveLicenses ($LicenseData | ForEach-Object { $_.SkuId }) -ErrorAction Stop
                $StatusMsg = "Todas as licencas removidas - $($DisplayName)"
                Write-Host "OK - " -ForegroundColor Green -NoNewline
                Write-Host $StatusMsg -ForegroundColor Green
                $RemovalReport.Add([PSCustomObject]@{
                    Numero                      = $i
                    Nome                        = $DisplayName
                    Email                       = $User.Mail
                    Departamento                = $User.Department
                    Licenca                     = "Todas"
                    "Data/Hora da execucao"     = (Get-Date -format "dd/MM/yyyy HH:mm:ss")
                })
            } catch {
                $ErrorMsg = "Erro ao remover todas as licencas do usuario ${User.Email}: $($_.Exception.Message)"
                Write-Host $ErrorMsg -ForegroundColor Red
                $RemovalReport.Add([PSCustomObject]@{
                    Numero                      = $i
                    Nome                        = $DisplayName
                    Email                       = $User.Mail
                    Departamento                = $User.Department
                    Licenca                     = "Todas"
                    "Data/Hora da execucao"     = (Get-Date -format "dd/MM/yyyy HH:mm:ss")
                    Erro                        = $ErrorMsg
                })
                Log-Error $ErrorMsg
            }
        } else {
            $StatusMsg = "Nenhuma licenca atribuida a conta de usuario ${User.Email}"
            Write-Host $StatusMsg -ForegroundColor Yellow
            $RemovalReport.Add([PSCustomObject]@{
                Numero                      = $i
                Nome                        = $DisplayName
                Email                       = $User.Mail
                Departamento                = $User.Department
                Licenca                     = "Nenhuma"
                "Data/Hora da execucao"     = (Get-Date -format "dd/MM/yyyy HH:mm:ss")
                Erro                        = $StatusMsg
            })
        }
    }
    return $RemovalReport
}

# Funcao para mostrar o menu
function Show-Menu {
    Write-Host "Selecione uma opcao:" -ForegroundColor Cyan
    Write-Host "1 - Adicionar Licencas" -ForegroundColor Yellow
    Write-Host "2 - Adicionar e Remover Licencas" -ForegroundColor Yellow
    Write-Host "3 - Remover Licencas" -ForegroundColor Yellow
    Write-Host "4 - Remover Todas as Licencas" -ForegroundColor Yellow
    Write-Host "0 - Sair" -ForegroundColor Yellow
    [int]$choice = Read-Host "Escolha uma opcao"
    return $choice
}

# Funcao para redimensionar a janela do PowerShell
function Resize-Window {
    $pshost = get-host
    $pswindow = $pshost.ui.rawui
    $newsize = $pswindow.buffersize
    $newsize.width = 120
    $newsize.height = 3000
    $pswindow.buffersize = $newsize
    $newsize = $pswindow.windowsize
    $newsize.width = 120
    $newsize.height = 50
    $pswindow.windowsize = $newsize
}

# Funcao principal para executar o script
function Main {
    Resize-Window
    Show-WelcomeScreen

    Check-Install-Modules

    try {
        Connect-MicrosoftGraph
    } catch {
        Write-Host "Erro ao conectar aos servicos. Saindo..." -ForegroundColor Red
        exit
    }

    try {
        while ($true) {
            $choice = Show-Menu

            switch ($choice) {
                1 {
                    Write-Host "Adicionar Licencas"
                    $users = Import-UsersFromCsv
                    if ($users) {
                        $skus = Get-AvailableSkus
                        Write-Host "Selecione a(s) licenca(s) a ser(em) atribuida(s) (digite os numeros separados por espacos):"
                        $i = 1
                        foreach ($sku in $skus) {
                            Write-Host "$i. $($sku.SkuPartNumber) - Disponivel: $($sku.Disponivel)"
                            $i++
                        }
                        $selectedSkuIndexes = (Read-Host "Escolha as licencas pelo numero: ").Split(" ") | ForEach-Object { [int]$_ }
                        $selectedSkus = $selectedSkuIndexes | ForEach-Object { $skus[$_ - 1].SkuId }
                        $report = Assign-Licenses -skuIds $selectedSkus -users $users
                        # Removendo a geração do relatório e a abertura do navegador
                    }
                }
                2 {
                    Write-Host "Adicionar e Remover Licencas"
                    $users = Import-UsersFromCsv
                    if ($users) {
                        $skus = Get-AvailableSkus
                        Write-Host "Selecione a(s) licenca(s) a ser(em) atribuida(s) (digite os numeros separados por espacos):"
                        $i = 1
                        foreach ($sku in $skus) {
                            Write-Host "$i. $($sku.SkuPartNumber) - Disponivel: $($sku.Disponivel)"
                            $i++
                        }
                        $selectedAddSkuIndexes = (Read-Host "Escolha as licencas pelo numero: ").Split(" ") | ForEach-Object { [int]$_ }
                        $selectedAddSkus = $selectedAddSkuIndexes | ForEach-Object { $skus[$_ - 1].SkuId }

                        Write-Host ""
                        Write-Host "Selecione a(s) licenca(s) a ser(em) removida(s) (digite os numeros separados por espacos):" -ForegroundColor Red
                        $i = 1
                        foreach ($sku in $skus) {
                            Write-Host "$i. $($sku.SkuPartNumber) - Disponivel: $($sku.Disponivel)"
                            $i++
                        }
                        $selectedRemoveSkuIndexes = (Read-Host "Escolha as licencas pelo numero: ").Split(" ") | ForEach-Object { [int]$_ }
                        $selectedRemoveSkus = $selectedRemoveSkuIndexes | ForEach-Object { $skus[$_ - 1].SkuId }

                        $assignReport = Assign-Licenses -skuIds $selectedAddSkus -users $users
                        $removeReport = Remove-Licenses -skuIds $selectedRemoveSkus -users $users
                        # Removendo a geração do relatório e a abertura do navegador
                    }
                }
                3 {
                    Write-Host "Remover Licencas"
                    $users = Import-UsersFromCsv
                    if ($users) {
                        $skus = Get-AvailableSkus
                        Write-Host "Selecione a(s) licenca(s) a ser(em) removida(s) (digite os numeros separados por espacos):"
                        $i = 1
                        foreach ($sku in $skus) {
                            Write-Host "$i. $($sku.SkuPartNumber) - Disponivel: $($sku.Disponivel)"
                            $i++
                        }
                        $selectedSkuIndexes = (Read-Host "Escolha as licencas pelo numero: ").Split(" ") | ForEach-Object { [int]$_ }
                        $selectedSkus = $selectedSkuIndexes | ForEach-Object { $skus[$_ - 1].SkuId }
                        $report = Remove-Licenses -skuIds $selectedSkus -users $users
                        # Removendo a geração do relatório e a abertura do navegador
                    }
                }
                4 {
                    Write-Host "Remover Todas as Licencas"
                    $users = Import-UsersFromCsv
                    if ($users) {
                        $report = Remove-AllLicenses -users $users
                        # Removendo a geração do relatório e a abertura do navegador
                    }
                }
                0 {
                    Write-Host "Saindo..."
                    break
                }
                default {
                    Write-Host "Opcao invalida. Tente novamente." -ForegroundColor Red
                }
            }
        }
    } finally {
        Disconnect-Services
    }
}

Main

# Criar arquivo CSV com emails ficticios
$csvContent = @"
Email,DisplayName,Department
ronaldo@exemplo.com,Ronaldo,Finance
debora@exemplo.com,Debora,HR
jose@exemplo.com,Jose,IT
maria@exemplo.com,Maria,Marketing
antonio@exemplo.com,Antonio,Sales
"@
$csvPath = "C:\MS365\contas.csv"
Set-Content -Path $csvPath -Value $csvContent -Force
