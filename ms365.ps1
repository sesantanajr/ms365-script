# Importa modulos necessarios
Import-Module AzureAD
Import-Module Microsoft.PowerShell.Management

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

# Funcao para exibir uma mensagem amigavel
function Show-Message {
    param (
        [string]$Message,
        [string]$Type = "Info"
    )

    if ($Type -eq "Error") {
        Write-Host "Erro: $Message" -ForegroundColor Red
        Log-Error $Message
    } elseif ($Type -eq "Attention") {
        Write-Host "Atencao: $Message" -ForegroundColor Yellow
        Log-Error $Message
    } else {
        Write-Host "$Message" -ForegroundColor Green
    }
}

# Funcao para exibir GUI de selecao de localidade
function Select-Location {
    Add-Type -AssemblyName System.Windows.Forms
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Selecione a Localidade Padrao"
    $form.Size = New-Object System.Drawing.Size(400, 200)
    $form.StartPosition = "CenterScreen"

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Localidade:"
    $label.AutoSize = $true
    $label.Top = 20
    $label.Left = 20
    $form.Controls.Add($label)

    $comboBox = New-Object System.Windows.Forms.ComboBox
    $comboBox.Width = 300
    $comboBox.Top = 50
    $comboBox.Left = 50
    $comboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $comboBox.Items.AddRange(@(
        "AF", "AL", "DZ", "AS", "AD", "AO", "AI", "AQ", "AG", "AR", "AM", "AW", "AU", "AT", "AZ", "BS", "BH", "BD", "BB", "BY", "BE", "BZ", "BJ", "BM", "BT", "BO", "BA", "BW", "BR", "BN", "BG", "BF", "BI", "KH", "CM", "CA", "CV", "KY", "CF", "TD", "CL", "CN", "CX", "CC", "CO", "KM", "CG", "CD", "CK", "CR", "CI", "HR", "CU", "CY", "CZ", "DK", "DJ", "DM", "DO", "EC", "EG", "SV", "GQ", "ER", "EE", "SZ", "ET", "FK", "FO", "FJ", "FI", "FR", "GF", "PF", "TF", "GA", "GM", "GE", "DE", "GH", "GI", "GR", "GL", "GD", "GP", "GU", "GT", "GG", "GN", "GW", "GY", "HT", "HM", "VA", "HN", "HK", "HU", "IS", "IN", "ID", "IR", "IQ", "IE", "IM", "IL", "IT", "JM", "JP", "JE", "JO", "KZ", "KE", "KI", "KP", "KR", "KW", "KG", "LA", "LV", "LB", "LS", "LR", "LY", "LI", "LT", "LU", "MO", "MK", "MG", "MW", "MY", "MV", "ML", "MT", "MH", "MQ", "MR", "MU", "YT", "MX", "FM", "MD", "MC", "MN", "ME", "MS", "MA", "MZ", "MM", "NA", "NR", "NP", "NL", "NC", "NZ", "NI", "NE", "NG", "NU", "NF", "MP", "NO", "OM", "PK", "PW", "PS", "PA", "PG", "PY", "PE", "PH", "PN", "PL", "PT", "PR", "QA", "RE", "RO", "RU", "RW", "BL", "SH", "KN", "LC", "MF", "PM", "VC", "WS", "SM", "ST", "SA", "SN", "RS", "SC", "SL", "SG", "SX", "SK", "SI", "SB", "SO", "ZA", "GS", "SS", "ES", "LK", "SD", "SR", "SJ", "SE", "CH", "SY", "TW", "TJ", "TZ", "TH", "TL", "TG", "TK", "TO", "TT", "TN", "TR", "TM", "TC", "TV", "UG", "UA", "AE", "GB", "US", "UM", "UY", "UZ", "VU", "VE", "VN", "VG", "VI", "WF", "EH", "YE", "ZM", "ZW"
    ))
    $comboBox.SelectedItem = "BR"
    $form.Controls.Add($comboBox)

    $buttonOk = New-Object System.Windows.Forms.Button
    $buttonOk.Text = "OK"
    $buttonOk.Top = 100
    $buttonOk.Left = 150
    $buttonOk.Add_Click({
        $form.Tag = $comboBox.SelectedItem
        $form.Close()
    })
    $form.Controls.Add($buttonOk)

    $form.ShowDialog()
    return $form.Tag
}

# Funcao para exibir GUI de selecao de arquivo CSV
function Select-CSVFile {
    Add-Type -AssemblyName System.Windows.Forms
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Filter = "CSV files (*.csv)|*.csv"
    $fileDialog.ShowDialog() | Out-Null
    return $fileDialog.FileName
}

# Funcao para verificar e instalar/atualizar modulo Microsoft Graph
function Check-Install-Modules {
    $moduleName = "Microsoft.Graph"
    try {
        if (-not (Get-Module -Name $moduleName -ListAvailable)) {
            Show-Message "Modulo $moduleName nao encontrado. Tentando instalar..." -Type "Info"
            Install-Module -Name $moduleName -Scope CurrentUser -Force -ErrorAction Stop
        } else {
            Show-Message "Modulo $moduleName ja esta instalado." -Type "Info"
        }
    } catch {
        $errorMessage = "Erro ao instalar/atualizar o modulo ${moduleName}: $($_.Exception.Message)"
        Show-Message $errorMessage -Type "Error"
    }
}

# Tela de boas-vindas
function Show-WelcomeScreen {
    Clear-Host
    Write-Host "===========================================" -ForegroundColor Green
    Write-Host "                                         " -ForegroundColor Green
    Write-Host "              JORNADA365                  " -ForegroundColor Green
    Write-Host "            Sua Jornada Comenca Aqui      " -ForegroundColor Green
    Write-Host "                                         " -ForegroundColor Green
    Write-Host "===========================================" -ForegroundColor Green
    Write-Host "                                         " -ForegroundColor Yellow
    Write-Host "Este script foi criado para simplificar o gerenciamento de licencas no Microsoft 365." -ForegroundColor Yellow
    Write-Host "                                         " -ForegroundColor Yellow
    Write-Host "Gerenciar licencas e uma tarefa facil quando se trata de poucas licencas ou mesmo" -ForegroundColor Yellow
    Write-Host "algumas dezenas de usuarios. No entanto, ao remover ou substituir licencas para" -ForegroundColor Yellow
    Write-Host "centenas de usuarios, o processo se torna muito mais complexo." -ForegroundColor Yellow
    Write-Host "Por isso, utilize este script com cautela e sinta-se a vontade para compartilha-lo." -ForegroundColor Yellow
    Write-Host "                                         " -ForegroundColor Yellow
    Write-Host "Visite nosso site: www.jornada365.cloud" -ForegroundColor Cyan
    Write-Host "Microsoft 365: admin.microsoft.com" -ForegroundColor Cyan
    Write-Host "                                         " -ForegroundColor Cyan
    Write-Host "SEJA BEM-VINDO || JORNADA 365" -ForegroundColor Green
    Write-Host "                                         " -ForegroundColor Green
    Write-Host "===========================================" -ForegroundColor Green
}

# Conectar ao Microsoft Graph com autenticacao interativa
function Connect-MicrosoftGraph {
    $maxRetries = 5
    $retryCount = 0
    $connected = $false

    while (-not $connected -and $retryCount -lt $maxRetries) {
        try {
            Show-Message "Conectando ao Microsoft Graph..." -Type "Info"
            Connect-MgGraph -Scopes "Directory.ReadWrite.All", "User.ReadWrite.All", "Group.ReadWrite.All", "Sites.ReadWrite.All", "DeviceManagementManagedDevices.ReadWrite.All", "Reports.Read.All", "Mail.ReadWrite"
            Show-Message "Conectado ao Microsoft Graph." -Type "Info"
            $connected = $true
        } catch {
            $retryCount++
            $errorMessage = "Erro de autenticacao: $($_.Exception.Message). Tentativa $retryCount de $maxRetries."
            Show-Message $errorMessage -Type "Error"
            Start-Sleep -Seconds 10
        }
    }

    if (-not $connected) {
        Show-Message "Falha ao conectar ao Microsoft Graph apos $maxRetries tentativas." -Type "Error"
        exit
    }
}

# Desconectar do Microsoft Graph
function Disconnect-Services {
    Show-Message "Desconectando do Microsoft Graph..." -Type "Info"
    try {
        Disconnect-MgGraph -ErrorAction Stop
        Show-Message "Desconectado do Microsoft Graph." -Type "Info"
    } catch {
        Show-Message "Nenhuma aplicacao para desconectar." -Type "Info"
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
        Show-Message $errorMessage -Type "Error"
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

    return $SkuList
}

# Importar usuarios de um arquivo CSV
function Import-UsersFromCsv {
    $filePath = Select-CSVFile

    if ($filePath) {
        $csvData = Import-Csv -Path $filePath
        $csvData | ForEach-Object {
            if (-not $_.PSObject.Properties["UserPrincipalName"]) {
                $_ | Add-Member -MemberType NoteProperty -Name UserPrincipalName -Value $_.Email
            }
        }
        return $csvData
    } else {
        Show-Message "Nenhum arquivo selecionado." -Type "Error"
        return $null
    }
}

# Verificar e definir localidade do usuario
function Check-Set-UserLocation {
    param (
        [array]$users,
        [string]$location
    )

    $LocationReport = [System.Collections.Generic.List[PSCustomObject]]::new()
    $i = 0
    foreach ($User in $Users) {
        if (-not $User.UserPrincipalName) {
            $notificationMessage = "Usuario nao possui um UPN valido. Pulando usuario."
            Show-Message $notificationMessage -Type "Attention"
            continue
        }

        $i++
        try {
            $UserData = Get-MgUser -UserId $User.UserPrincipalName.Trim() -Property id, usageLocation, displayName -ErrorAction Stop
            if (-not $UserData.UsageLocation) {
                Set-MgUser -UserId $User.UserPrincipalName.Trim() -UsageLocation $location
                $LocationStatus = "Localidade definida para $location"
                Show-Message "$LocationStatus para $($UserData.DisplayName)" -Type "Info"
            } else {
                $LocationStatus = "Localidade ja definida ($($UserData.UsageLocation))"
                Show-Message "$LocationStatus para $($UserData.DisplayName)" -Type "Info"
            }
            $LocationReport.Add([PSCustomObject]@{
                Numero = $i
                Nome = $UserData.DisplayName
                UPN = $User.UserPrincipalName
                Localidade = if ($UserData.UsageLocation) { $UserData.UsageLocation } else { $location }
                Status = $LocationStatus
            })
        } catch {
            if ($_.Exception.ErrorCode -eq "Request_ResourceNotFound") {
                $notificationMessage = "Usuario ${User.UserPrincipalName} nao existe ou foi excluido."
                Show-Message $notificationMessage -Type "Attention"
            } else {
                $errorMessage = "Erro ao definir localidade do usuario ${User.UserPrincipalName}: $($_.Exception.Message)"
                Show-Message $errorMessage -Type "Error"
            }
        }
    }
    return $LocationReport
}

# Atribuir licencas
function Assign-Licenses {
    param (
        [array]$skuIds,
        [array]$users,
        [string]$defaultLocation
    )

    $AssignmentReport = [System.Collections.Generic.List[PSCustomObject]]::new()
    $i = 0
    $friendlySkuNames = Get-FriendlySkuNames
    foreach ($User in $Users) {
        if (-not $User.UserPrincipalName) {
            $notificationMessage = "Usuario nao possui um UPN valido. Pulando usuario."
            Show-Message $notificationMessage -Type "Attention"
            continue
        }

        $AttentionMsg = $Null; $i++
        Show-Message ("Processando conta de $i/$($Users.Count)") -Type "Info"
        try {
            $UserData = Get-MgUser -UserId $User.UserPrincipalName.Trim() -Property id, assignedLicenses, department, displayName, usageLocation -ErrorAction Stop
            if (-not $UserData.UsageLocation) {
                Show-Message "Localidade nao definida para $($UserData.DisplayName). Definindo localidade padrao..." -Type "Info"
                Set-MgUser -UserId $User.UserPrincipalName.Trim() -UsageLocation $defaultLocation
            }
            $DisplayName = $UserData.DisplayName
        } catch {
            if ($_.Exception.ErrorCode -eq "Request_ResourceNotFound") {
                $notificationMessage = "Usuario ${User.UserPrincipalName} nao existe ou foi excluido."
                Show-Message $notificationMessage -Type "Attention"
            } else {
                $AttentionMsg = "Erro ao buscar dados do usuario ${User.UserPrincipalName}: $($_.Exception.Message)"
                Show-Message $AttentionMsg -Type "Error"
                $AssignmentReport.Add([PSCustomObject]@{
                    Numero                      = $i
                    Nome                        = $User.DisplayName
                    UPN                         = $User.UserPrincipalName
                    Departamento                = $User.Department
                    Licenca                     = $friendlySkuNames[$skuId]
                    "Data/Hora da execucao"     = (Get-Date -format "dd/MM/yyyy HH:mm:ss")
                    Erro                        = $AttentionMsg
                })
            }
            continue
        }

        foreach ($skuId in $skuIds) {
            $LicenseData = $UserData | Select-Object -ExpandProperty AssignedLicenses
            if ($skuId -in $LicenseData.SkuId) {
                $StatusMsg = "Licenca ja atribuida a conta de usuario ${User.UserPrincipalName}"
                Show-Message $StatusMsg -Type "Info"
                $AssignmentReport.Add([PSCustomObject]@{
                    Numero                      = $i
                    Nome                        = $DisplayName
                    UPN                         = $User.UserPrincipalName
                    Departamento                = $User.Department
                    Licenca                     = $friendlySkuNames[$skuId]
                    "Data/Hora da execucao"     = (Get-Date -format "dd/MM/yyyy HH:mm:ss")
                    Status                      = $StatusMsg
                })
            } else {
                try {
                    Set-MgUserLicense -UserId $User.UserPrincipalName -Addlicenses @{SkuId = $skuId} -RemoveLicenses @() -ErrorAction Stop
                    $StatusMsg = "Licenca atribuida $($friendlySkuNames[$skuId]) - $($DisplayName)"
                    Show-Message $StatusMsg -Type "Info"
                    $AssignmentReport.Add([PSCustomObject]@{
                        Numero                      = $i
                        Nome                        = $DisplayName
                        UPN                         = $User.UserPrincipalName
                        Departamento                = $User.Department
                        Licenca                     = $friendlySkuNames[$skuId]
                        "Data/Hora da execucao"     = (Get-Date -format "dd/MM/yyyy HH:mm:ss")
                        Status                      = $StatusMsg
                    })
                } catch {
                    if ($_.Exception.ErrorCode -eq "Request_BadRequest" -and $_.Exception.Message -match "does not have any available licenses") {
                        $InfoMsg = "Nao ha licencas disponiveis para atribuir a conta de usuario ${User.UserPrincipalName}."
                        Show-Message $InfoMsg -Type "Attention"
                        $AssignmentReport.Add([PSCustomObject]@{
                            Numero                      = $i
                            Nome                        = $DisplayName
                            UPN                         = $User.UserPrincipalName
                            Departamento                = $User.Department
                            Licenca                     = $friendlySkuNames[$skuId]
                            "Data/Hora da execucao"     = (Get-Date -format "dd/MM/yyyy HH:mm:ss")
                            Status                      = $InfoMsg
                        })
                    } elseif ($_.Exception.Message -match "mutually exclusive") {
                        $InfoMsg = "Nao foi possivel atribuir a licenca ao usuario ${User.UserPrincipalName} pois ha conflito entre planos de servico."
                        Show-Message $InfoMsg -Type "Attention"
                        $AssignmentReport.Add([PSCustomObject]@{
                            Numero                      = $i
                            Nome                        = $DisplayName
                            UPN                         = $User.UserPrincipalName
                            Departamento                = $User.Department
                            Licenca                     = $friendlySkuNames[$skuId]
                            "Data/Hora da execucao"     = (Get-Date -format "dd/MM/yyyy HH:mm:ss")
                            Status                      = $InfoMsg
                        })
                    } else {
                        $AttentionMsg = "Nao foi possivel atribuir a licenca ao usuario ${User.UserPrincipalName}: $($_.Exception.Message)"
                        Show-Message $AttentionMsg -Type "Error"
                        $AssignmentReport.Add([PSCustomObject]@{
                            Numero                      = $i
                            Nome                        = $DisplayName
                            UPN                         = $User.UserPrincipalName
                            Departamento                = $User.Department
                            Licenca                     = $friendlySkuNames[$skuId]
                            "Data/Hora da execucao"     = (Get-Date -format "dd/MM/yyyy HH:mm:ss")
                            Erro                        = $AttentionMsg
                        })
                    }
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
        if (-not $User.UserPrincipalName) {
            $notificationMessage = "Usuario nao possui um UPN valido. Pulando usuario."
            Show-Message $notificationMessage -Type "Attention"
            continue
        }

        $AttentionMsg = $Null; $i++
        Show-Message ("Processando conta de $i/$($Users.Count)") -Type "Info"
        try {
            $UserData = Get-MgUser -UserId $User.UserPrincipalName.Trim() -Property id, assignedLicenses, department, displayName -ErrorAction Stop
            $DisplayName = $UserData.DisplayName
        } catch {
            if ($_.Exception.ErrorCode -eq "Request_ResourceNotFound") {
                $notificationMessage = "Usuario ${User.UserPrincipalName} nao existe ou foi excluido."
                Show-Message $notificationMessage -Type "Attention"
            } else {
                $AttentionMsg = "Erro ao buscar dados do usuario ${User.UserPrincipalName}: $($_.Exception.Message)"
                Show-Message $AttentionMsg -Type "Error"
                $RemovalReport.Add([PSCustomObject]@{
                    Numero                      = $i
                    Nome                        = $User.DisplayName
                    UPN                         = $User.UserPrincipalName
                    Departamento                = $User.Department
                    Licenca                     = $friendlySkuNames[$skuId]
                    "Data/Hora da execucao"     = (Get-Date -format "dd/MM/yyyy HH:mm:ss")
                    Erro                        = $AttentionMsg
                })
            }
            continue
        }

        foreach ($skuId in $skuIds) {
            $LicenseData = $UserData | Select-Object -ExpandProperty AssignedLicenses
            if ($skuId -in $LicenseData.SkuId) {
                try {
                    Set-MgUserLicense -UserId $User.UserPrincipalName -Addlicenses @() -RemoveLicenses @($skuId) -ErrorAction Stop
                    $StatusMsg = "Licenca removida $($friendlySkuNames[$skuId]) - $($DisplayName)"
                    Show-Message $StatusMsg -Type "Info"
                    $RemovalReport.Add([PSCustomObject]@{
                        Numero                      = $i
                        Nome                        = $DisplayName
                        UPN                         = $User.UserPrincipalName
                        Departamento                = $User.Department
                        Licenca                     = $friendlySkuNames[$skuId]
                        "Data/Hora da execucao"     = (Get-Date -format "dd/MM/yyyy HH:mm:ss")
                        Status                      = $StatusMsg
                    })
                } catch {
                    $AttentionMsg = "Nao foi possivel remover a licenca do usuario ${User.UserPrincipalName}: $($_.Exception.Message)"
                    Show-Message $AttentionMsg -Type "Error"
                    $RemovalReport.Add([PSCustomObject]@{
                        Numero                      = $i
                        Nome                        = $DisplayName
                        UPN                         = $User.UserPrincipalName
                        Departamento                = $User.Department
                        Licenca                     = $friendlySkuNames[$skuId]
                        "Data/Hora da execucao"     = (Get-Date -format "dd/MM/yyyy HH:mm:ss")
                        Erro                        = $AttentionMsg
                    })
                }
            } else {
                $StatusMsg = "Licenca nao atribuida a conta de usuario ${User.UserPrincipalName}"
                Show-Message $StatusMsg -Type "Info"
                $RemovalReport.Add([PSCustomObject]@{
                    Numero                      = $i
                    Nome                        = $DisplayName
                    UPN                         = $User.UserPrincipalName
                    Departamento                = $User.Department
                    Licenca                     = $friendlySkuNames[$skuId]
                    "Data/Hora da execucao"     = (Get-Date -format "dd/MM/yyyy HH:mm:ss")
                    Status                      = $StatusMsg
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
        if (-not $User.UserPrincipalName) {
            $notificationMessage = "Usuario nao possui um UPN valido. Pulando usuario."
            Show-Message $notificationMessage -Type "Attention"
            continue
        }

        $AttentionMsg = $Null; $i++
        Show-Message ("Processando conta de $i/$($Users.Count)") -Type "Info"
        try {
            $UserData = Get-MgUser -UserId $User.UserPrincipalName.Trim() -Property id, assignedLicenses, department, displayName -ErrorAction Stop
            $DisplayName = $UserData.DisplayName
        } catch {
            if ($_.Exception.ErrorCode -eq "Request_ResourceNotFound") {
                $notificationMessage = "Usuario ${User.UserPrincipalName} nao existe ou foi excluido."
                Show-Message $notificationMessage -Type "Attention"
            } else {
                $AttentionMsg = "Erro ao buscar dados do usuario ${User.UserPrincipalName}: $($_.Exception.Message)"
                Show-Message $AttentionMsg -Type "Error"
                $RemovalReport.Add([PSCustomObject]@{
                    Numero                      = $i
                    Nome                        = $User.DisplayName
                    UPN                         = $User.UserPrincipalName
                    Departamento                = $User.Department
                    Licenca                     = "Todas"
                    "Data/Hora da execucao"     = (Get-Date -format "dd/MM/yyyy HH:mm:ss")
                    Erro                        = $AttentionMsg
                })
            }
            continue
        }

        $LicenseData = $UserData | Select-Object -ExpandProperty AssignedLicenses
        if ($LicenseData.Count -gt 0) {
            try {
                Set-MgUserLicense -UserId $User.UserPrincipalName -Addlicenses @() -RemoveLicenses ($LicenseData | ForEach-Object { $_.SkuId }) -ErrorAction Stop
                $StatusMsg = "Todas as licencas removidas - $($DisplayName)"
                Show-Message $StatusMsg -Type "Info"
                $RemovalReport.Add([PSCustomObject]@{
                    Numero                      = $i
                    Nome                        = $DisplayName
                    UPN                         = $User.UserPrincipalName
                    Departamento                = $User.Department
                    Licenca                     = "Todas"
                    "Data/Hora da execucao"     = (Get-Date -format "dd/MM/yyyy HH:mm:ss")
                    Status                      = $StatusMsg
                })
            } catch {
                $AttentionMsg = "Nao foi possivel remover todas as licencas do usuario ${User.UserPrincipalName}: $($_.Exception.Message)"
                Show-Message $AttentionMsg -Type "Error"
                $RemovalReport.Add([PSCustomObject]@{
                    Numero                      = $i
                    Nome                        = $DisplayName
                    UPN                         = $User.UserPrincipalName
                    Departamento                = $User.Department
                    Licenca                     = "Todas"
                    "Data/Hora da execucao"     = (Get-Date -format "dd/MM/yyyy HH:mm:ss")
                    Erro                        = $AttentionMsg
                })
            }
        } else {
            $StatusMsg = "Nenhuma licenca atribuida a conta de usuario ${User.UserPrincipalName}"
            Show-Message $StatusMsg -Type "Info"
            $RemovalReport.Add([PSCustomObject]@{
                Numero                      = $i
                Nome                        = $DisplayName
                UPN                         = $User.UserPrincipalName
                Departamento                = $User.Department
                Licenca                     = "Nenhuma"
                "Data/Hora da execucao"     = (Get-Date -format "dd/MM/yyyy HH:mm:ss")
                Status                      = $StatusMsg
            })
        }
    }
    return $RemovalReport
}

# Funcao para mostrar o menu
function Show-Menu {
    Write-Host " " -ForegroundColor Cyan
    Write-Host "Selecione uma opcao:" -ForegroundColor Cyan
    Write-Host "1 - Adicionar Licencas" -ForegroundColor Yellow
    Write-Host "2 - Adicionar e Remover Licencas" -ForegroundColor Yellow
    Write-Host "3 - Remover Licencas" -ForegroundColor Yellow
    Write-Host "4 - Remover Todas as Licencas" -ForegroundColor Yellow
    Write-Host "5 - Definir Localidade Padrao para Todos os Usuarios" -ForegroundColor Yellow
    Write-Host "6 - Importar CSV e Definir Localidade" -ForegroundColor Yellow
    Write-Host "0 - Sair" -ForegroundColor Yellow
    Write-Host " " -ForegroundColor Yellow
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
        Show-Message "Erro ao conectar aos servicos. Saindo..." -Type "Error"
        exit
    }

    try {
        while ($true) {
            $choice = Show-Menu

            switch ($choice) {
                1 {
                    Show-Message "Adicionar Licencas" -Type "Info"
                    $users = Import-UsersFromCsv
                    if ($users) {
                        $skus = Get-AvailableSkus
                        Show-Message "Selecione a(s) licenca(s) a ser(em) atribuida(s) (digite os numeros separados por espacos):" -Type "Info"
                        $i = 1
                        foreach ($sku in $skus) {
                            Write-Host "$i. $($sku.SkuPartNumber) - Disponivel: $($sku.Disponivel)"
                            $i++
                        }
                        $selectedSkuIndexes = (Read-Host "Escolha as licencas pelo numero: ").Split(" ") | ForEach-Object { [int]$_ }
                        $selectedSkus = $selectedSkuIndexes | ForEach-Object { $skus[$_ - 1].SkuId }
                        $defaultLocation = Select-Location
                        $report = Assign-Licenses -skuIds $selectedSkus -users $users -defaultLocation $defaultLocation
                    }
                }
                2 {
                    Show-Message "Adicionar e Remover Licencas" -Type "Info"
                    $users = Import-UsersFromCsv
                    if ($users) {
                        $skus = Get-AvailableSkus
                        Show-Message "Selecione a(s) licenca(s) a ser(em) atribuida(s) (digite os numeros separados por espacos):" -Type "Info"
                        $i = 1
                        foreach ($sku in $skus) {
                            Write-Host "$i. $($sku.SkuPartNumber) - Disponivel: $($sku.Disponivel)"
                            $i++
                        }
                        $selectedAddSkuIndexes = (Read-Host "Escolha as licencas pelo numero: ").Split(" ") | ForEach-Object { [int]$_ }
                        $selectedAddSkus = $selectedAddSkuIndexes | ForEach-Object { $skus[$_ - 1].SkuId }

                        Write-Host ""
                        Show-Message "Selecione a(s) licenca(s) a ser(em) removida(s) (digite os numeros separados por espacos):" -Type "Info"
                        $i = 1
                        foreach ($sku in $skus) {
                            Write-Host "$i. $($sku.SkuPartNumber) - Disponivel: $($sku.Disponivel)"
                            $i++
                        }
                        $selectedRemoveSkuIndexes = (Read-Host "Escolha as licencas pelo numero: ").Split(" ") | ForEach-Object { [int]$_ }
                        $selectedRemoveSkus = $selectedRemoveSkuIndexes | ForEach-Object { $skus[$_ - 1].SkuId }

                        $defaultLocation = Select-Location
                        $assignReport = Assign-Licenses -skuIds $selectedAddSkus -users $users -defaultLocation $defaultLocation
                        $removeReport = Remove-Licenses -skuIds $selectedRemoveSkus -users $users
                    }
                }
                3 {
                    Show-Message "Remover Licencas" -Type "Info"
                    $users = Import-UsersFromCsv
                    if ($users) {
                        $skus = Get-AvailableSkus
                        Show-Message "Selecione a(s) licenca(s) a ser(em) removida(s) (digite os numeros separados por espacos):" -Type "Info"
                        $i = 1
                        foreach ($sku in $skus) {
                            Write-Host "$i. $($sku.SkuPartNumber) - Disponivel: $($sku.Disponivel)"
                            $i++
                        }
                        $selectedSkuIndexes = (Read-Host "Escolha as licencas pelo numero: ").Split(" ") | ForEach-Object { [int]$_ }
                        $selectedSkus = $selectedSkuIndexes | ForEach-Object { $skus[$_ - 1].SkuId }
                        $report = Remove-Licenses -skuIds $selectedSkus -users $users
                    }
                }
                4 {
                    Show-Message "Remover Todas as Licencas" -Type "Info"
                    $users = Import-UsersFromCsv
                    if ($users) {
                        $report = Remove-AllLicenses -users $users
                    }
                }
                5 {
                    Show-Message "Definir Localidade Padrao para Todos os Usuarios" -Type "Info"
                    $defaultLocation = Select-Location
                    if ($defaultLocation) {
                        $users = Get-MgUser -All
                        if ($users) {
                            $report = Check-Set-UserLocation -users $users -location $defaultLocation
                        } else {
                            Show-Message "Nenhum usuario encontrado." -Type "Error"
                        }
                    } else {
                        Show-Message "Nenhuma localidade selecionada. Operacao cancelada." -Type "Error"
                    }
                }
                6 {
                    Show-Message "Importar CSV e Definir Localidade" -Type "Info"
                    $users = Import-UsersFromCsv
                    if ($users) {
                        $defaultLocation = Select-Location
                        if ($defaultLocation) {
                            $report = Check-Set-UserLocation -users $users -location $defaultLocation
                        } else {
                            Show-Message "Nenhuma localidade selecionada. Operacao cancelada." -Type "Error"
                        }
                    }
                }
                0 {
                    Show-Message "Saindo..." -Type "Info"
                    break
                }
                default {
                    Show-Message "Opcao invalida. Tente novamente." -Type "Error"
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
Email
adelev@jornada365.cloud
aline.fonseca@jornada365.cloud
amauri.gomes@jornada365.cloud
andresa.fontes@jornada365.cloud
bete.luma@jornada365.cloud
"@
$csvPath = "C:\MS365\contas.csv"
Set-Content -Path $csvPath -Value $csvContent -Force
