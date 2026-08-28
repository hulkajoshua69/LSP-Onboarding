<#
.SYNOPSIS
GUI Tool for automating O365 user creation for Legacy Service Partners (LSP).

.DESCRIPTION
Reads configuration data from \Config\Companies.json. Connects to Microsoft Graph 
and Exchange Online to provision the account, set managers, apply explicit 
group-based licensing, and configure Exchange aliases.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path $script:MyInvocation.MyCommand.Path }
$JsonPath = Join-Path $ScriptDir "Config\Companies.json"

if (-not (Test-Path $JsonPath)) {
    [System.Windows.Forms.MessageBox]::Show("Configuration file missing! Please ensure Config\Companies.json exists.", "Error", 0, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}

$CompanyRef = Get-Content $JsonPath -Raw | ConvertFrom-Json
$CompanyNames = $CompanyRef.PSObject.Properties.Name | Sort-Object

$form = New-Object System.Windows.Forms.Form
$form.Text = "LSP Onboarding Automator"
$form.Size = New-Object System.Drawing.Size(550, 650)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

$fontRegular = New-Object System.Drawing.Font("Segoe UI", 9.5)
$fontBold = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)

function New-Label($text, $y, $x=20) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $text
    $lbl.Location = New-Object System.Drawing.Point($x, $y)
    $lbl.AutoSize = $true
    $lbl.Font = $fontRegular
    $form.Controls.Add($lbl)
}

New-Label "First Name:" 20
$txtFirstName = New-Object System.Windows.Forms.TextBox
$txtFirstName.Location = New-Object System.Drawing.Point(140, 18)
$txtFirstName.Size = New-Object System.Drawing.Size(350, 25)
$form.Controls.Add($txtFirstName)

New-Label "Last Name:" 60
$txtLastName = New-Object System.Windows.Forms.TextBox
$txtLastName.Location = New-Object System.Drawing.Point(140, 58)
$txtLastName.Size = New-Object System.Drawing.Size(350, 25)
$form.Controls.Add($txtLastName)

New-Label "Company:" 100
$cmbCompany = New-Object System.Windows.Forms.ComboBox
$cmbCompany.Location = New-Object System.Drawing.Point(140, 98)
$cmbCompany.Size = New-Object System.Drawing.Size(350, 25)
$cmbCompany.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$CompanyNames | ForEach-Object { [void]$cmbCompany.Items.Add($_) }
$form.Controls.Add($cmbCompany)

New-Label "Job Title:" 140
$txtJobTitle = New-Object System.Windows.Forms.TextBox
$txtJobTitle.Location = New-Object System.Drawing.Point(140, 138)
$txtJobTitle.Size = New-Object System.Drawing.Size(350, 25)
$form.Controls.Add($txtJobTitle)

New-Label "Manager Email:" 180
$txtManagerEmail = New-Object System.Windows.Forms.TextBox
$txtManagerEmail.Location = New-Object System.Drawing.Point(140, 178)
$txtManagerEmail.Size = New-Object System.Drawing.Size(350, 25)
$form.Controls.Add($txtManagerEmail)

New-Label "License Group:" 220
$cmbLicense = New-Object System.Windows.Forms.ComboBox
$cmbLicense.Location = New-Object System.Drawing.Point(140, 218)
$cmbLicense.Size = New-Object System.Drawing.Size(350, 25)
$cmbLicense.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void]$cmbLicense.Items.Add("AZ_E3 License")
[void]$cmbLicense.Items.Add("AZ_E1 License")
$cmbLicense.SelectedIndex = 0
$form.Controls.Add($cmbLicense)

$btnSubmit = New-Object System.Windows.Forms.Button
$btnSubmit.Location = New-Object System.Drawing.Point(140, 260)
$btnSubmit.Size = New-Object System.Drawing.Size(350, 40)
$btnSubmit.Text = "Provision New User"
$btnSubmit.Font = $fontBold
$btnSubmit.BackColor = [System.Drawing.Color]::LightGreen
$form.Controls.Add($btnSubmit)

$txtOutput = New-Object System.Windows.Forms.RichTextBox
$txtOutput.Location = New-Object System.Drawing.Point(20, 320)
$txtOutput.Size = New-Object System.Drawing.Size(490, 260)
$txtOutput.ReadOnly = $true
$txtOutput.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtOutput)

function Write-GuiLog($message) {
    $timestamp = (Get-Date).ToString("HH:mm:ss")
    $txtOutput.AppendText("[$timestamp] $message`r`n")
    $txtOutput.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

$btnSubmit.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtFirstName.Text) -or [string]::IsNullOrWhiteSpace($txtLastName.Text) -or [string]::IsNullOrWhiteSpace($cmbCompany.Text)) {
        [System.Windows.Forms.MessageBox]::Show("First Name, Last Name, and Company are required fields.", "Validation Error", 0, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $btnSubmit.Enabled = $false
    $txtOutput.Clear()

    $FirstName = $txtFirstName.Text.Trim()
    $LastName = $txtLastName.Text.Trim()
    $CompanyKey = $cmbCompany.Text
    $JobTitle = $txtJobTitle.Text.Trim()
    $ManagerEmail = $txtManagerEmail.Text.Trim()
    $LicenseGroupName = $cmbLicense.Text

    $targetCompany = $CompanyRef.$CompanyKey

    try {
        Write-GuiLog "Starting automation for $FirstName $LastName..."
        Write-GuiLog "IMPORTANT: The UI may freeze momentarily while Microsoft authentication prompts appear."
        
        Write-GuiLog "Connecting to Microsoft Graph..."
        Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.ReadWrite.All", "GroupMember.ReadWrite.All" -NoWelcome
        
        Write-GuiLog "Connecting to Exchange Online..."
        Connect-ExchangeOnline -ShowBanner:$false

        $FirstNameLower = $FirstName.ToLower()
        $LastNameLower = $LastName.ToLower()
        
        $BaseUPNPrefix = "$FirstNameLower.$LastNameLower"
        $TenantDomain = "legacyservicepartners.com"
        $UPN = "$BaseUPNPrefix@$TenantDomain"

        $existingUser = Get-MgUser -Filter "userPrincipalName eq '$UPN'" -ErrorAction SilentlyContinue
        if ($existingUser) {
            Write-GuiLog "Duplicate user found! Appending '01' to UPN."
            $BaseUPNPrefix = "${BaseUPNPrefix}01"
            $UPN = "$BaseUPNPrefix@$TenantDomain"
        }

        $FirstInitialLower = $FirstNameLower.Substring(0, 1)
        $LastInitialLower = $LastNameLower.Substring(0, 1)
        $PartnerPrefix = ""
        switch ($targetCompany.EmailFormat) {
            "FirstInitialLastName" { $PartnerPrefix = "$FirstInitialLower$LastNameLower" }
            "FirstName.LastName"   { $PartnerPrefix = "$FirstNameLower.$LastNameLower" }
            "FirstNameLastInitial" { $PartnerPrefix = "$FirstNameLower$LastInitialLower" }
            "FirstName"            { $PartnerPrefix = "$FirstNameLower" }
        }
        
        if ($existingUser) { $PartnerPrefix = "${PartnerPrefix}01" }
        $PartnerPrimaryEmail = "$PartnerPrefix@$($targetCompany.PrimaryDomain)"

        $FirstInitCap = $FirstName.Substring(0, 1).ToUpper()
        $LastInitCap = $LastName.Substring(0, 1).ToUpper()
        $InitialPassword = "${FirstInitCap}${LastInitCap}sp123451"

        Write-GuiLog "Creating Entra ID user: $UPN"
        $PasswordProfile = @{ Password = $InitialPassword; ForceChangePasswordNextSignIn = $true }
        $newUserParams = @{
            AccountEnabled = $true
            DisplayName = "$FirstName $LastName"
            MailNickname = $BaseUPNPrefix
            UserPrincipalName = $UPN
            GivenName = $FirstName
            Surname = $LastName
            JobTitle = $JobTitle
            CompanyName = $targetCompany.CompanyName 
            UsageLocation = "US"
            PasswordProfile = $PasswordProfile
        }
        $newUser = New-MgUser @newUserParams

        if (-not $newUser) { throw "Failed to create Entra ID user." }
        Write-GuiLog "User created successfully! Object ID: $($newUser.Id)"

        if ($ManagerEmail) {
            Write-GuiLog "Locating manager: $ManagerEmail..."
            $ManagerObj = Get-MgUser -ConsistencyLevel eventual -Search "`"mail:$ManagerEmail`"" -ErrorAction SilentlyContinue | Select-Object -First 1
            if (-not $ManagerObj) { $ManagerObj = Get-MgUser -Filter "userPrincipalName eq '$ManagerEmail'" -ErrorAction SilentlyContinue }

            if ($ManagerObj) {
                Set-MgUserManagerByRef -UserId $newUser.Id -ManagerId $ManagerObj.Id
                Write-GuiLog "Manager linked successfully."
            } else {
                Write-GuiLog "WARNING: Manager not found in Entra ID. Assign manually."
            }
        }

        Write-GuiLog "Assigning user to explicit license group: $LicenseGroupName..."
        $LicenseGroup = Get-MgGroup -Filter "DisplayName eq '$LicenseGroupName'" -ErrorAction SilentlyContinue
        if ($LicenseGroup) {
            New-MgGroupMember -GroupId $LicenseGroup.Id -DirectoryObjectId $newUser.Id
            Write-GuiLog "Assigned to $LicenseGroupName successfully."
        } else {
            Write-GuiLog "WARNING: Could not find Group '$LicenseGroupName'."
        }

        Write-GuiLog "Waiting for Exchange Online mailbox provisioning (Max 3 mins)..."
        $mailboxReady = $false
        $retryCount = 0
        while (-not $mailboxReady -and $retryCount -lt 12) {
            Start-Sleep -Seconds 15
            $checkMailbox = Get-Mailbox -Identity $UPN -ErrorAction SilentlyContinue
            if ($checkMailbox) { $mailboxReady = $true } 
            else { $retryCount++; Write-GuiLog "Still waiting for Exchange sync (Attempt $retryCount/12)..." }
        }

        if ($mailboxReady) {
            Write-GuiLog "Mailbox synced! Applying email aliases..."
            $EmailAddresses = @("SMTP:$PartnerPrimaryEmail", "smtp:$UPN")
            foreach ($alias in $targetCompany.AliasDomains) {
                $EmailAddresses += "smtp:$PartnerPrefix@$alias"
            }
            Set-Mailbox -Identity $UPN -EmailAddresses $EmailAddresses
            Write-GuiLog "Aliases applied successfully."
        } else {
            Write-GuiLog "WARNING: Exchange timeout. Manually set aliases in EAC."
        }

        Write-GuiLog "Disconnecting from Exchange..."
        Disconnect-ExchangeOnline -Confirm:$false

        $txtOutput.AppendText("`r`n===============================================`r`n")
        $txtOutput.AppendText("ONBOARDING TICKET RESOLUTION WORK NOTES:`r`n")
        $txtOutput.AppendText("===============================================`r`n")
        $txtOutput.AppendText("Username: $UPN`r`n")
        $txtOutput.AppendText("Password: $InitialPassword`r`n")
        $txtOutput.AppendText("Company Name: $($targetCompany.CompanyName)`r`n")
        $txtOutput.AppendText("Primary Email Address: $PartnerPrimaryEmail`r`n")
        $aliasList = if ($targetCompany.AliasDomains) { $targetCompany.AliasDomains -join ", " } else { "None" }
        $txtOutput.AppendText("Alias Email Addresses: $aliasList`r`n")
        $txtOutput.AppendText("License Assigned: Group-Based ($LicenseGroupName)`r`n`r`n")
        $txtOutput.AppendText("Process Complete. Confirmation: Followed KBA10470957, KBA10471107, and KBA10470953 for account creation, dynamic groups alignment, and alias routing.`r`n")
        $txtOutput.AppendText("===============================================`r`n")
        $txtOutput.ScrollToCaret()

    } catch {
        Write-GuiLog "ERROR: $($_.Exception.Message)"
    } finally {
        $btnSubmit.Enabled = $true
    }
})

[void]$form.ShowDialog()
