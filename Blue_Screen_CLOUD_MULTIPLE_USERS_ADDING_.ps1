# Install module if needed (uncomment to run once)
# Install-Module AzureAD -Scope CurrentUser -Force

# Load AzureAD Module (corrected)
Import-Module AzureAD -Force

# Add GUI components
Add-Type -AssemblyName System.Windows.Forms

# Create Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Azure AD Group Membership - Add Users"
$form.Size = New-Object System.Drawing.Size(500, 400)
$form.BackColor = 'Blue'

# Group Name Label & TextBox
$labelGroup = New-Object System.Windows.Forms.Label
$labelGroup.Text = "Enter Azure AD Group Display Name:"
$labelGroup.ForeColor = 'White'
$labelGroup.Location = New-Object System.Drawing.Point(10, 10)
$labelGroup.AutoSize = $true
$form.Controls.Add($labelGroup)

$textGroup = New-Object System.Windows.Forms.TextBox
$textGroup.Size = New-Object System.Drawing.Size(460, 20)
$textGroup.Location = New-Object System.Drawing.Point(10, 30)
$form.Controls.Add($textGroup)

# User List Label & Multiline TextBox
$labelUsers = New-Object System.Windows.Forms.Label
$labelUsers.Text = "Enter user email addresses (one per line):"
$labelUsers.ForeColor = 'White'
$labelUsers.Location = New-Object System.Drawing.Point(10, 65)
$labelUsers.AutoSize = $true
$form.Controls.Add($labelUsers)

$textUsers = New-Object System.Windows.Forms.TextBox
$textUsers.Multiline = $true
$textUsers.ScrollBars = 'Vertical'
$textUsers.Size = New-Object System.Drawing.Size(460, 200)
$textUsers.Location = New-Object System.Drawing.Point(10, 85)
$form.Controls.Add($textUsers)

# OK Button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "Add Users to Group"
$okButton.Location = New-Object System.Drawing.Point(180, 300)
$okButton.Add_Click({ $form.Close() })
$form.Controls.Add($okButton)

# Show Form
$form.ShowDialog()

# Read Inputs
$groupDisplayName = $textGroup.Text.Trim()
$userList = $textUsers.Text -split "`r`n" | Where-Object { $_.Trim() -ne "" }

# --- Connect to Azure AD ---
Write-Host "`n🔄 Connecting to Azure AD..." -ForegroundColor Cyan
if ([string]::IsNullOrWhiteSpace($adminUPN)) {
    Connect-AzureAD
} else {
    Connect-AzureAD -AccountId $adminUPN
}

# Fetch Group by Display Name
$group = Get-AzureADGroup -Filter "DisplayName eq '$groupDisplayName'"

if (-not $group) {
    Write-Host "❌ Group not found: $groupDisplayName" -ForegroundColor Red
    exit
}

# Add each user
foreach ($userEmail in $userList) {
    $user = Get-AzureADUser -Filter "UserPrincipalName eq '$userEmail'" -ErrorAction SilentlyContinue

    if (-not $user) {
        Write-Host "⚠️ User not found: $userEmail" -ForegroundColor Yellow
    } else {
        try {
            Add-AzureADGroupMember -ObjectId $group.ObjectId -RefObjectId $user.ObjectId
            Write-Host "✅ Successfully added: $userEmail" -ForegroundColor Green
        } catch {
            Write-Host "❌ Error adding ${userEmail}: $_" -ForegroundColor Red
        }
    }
}

# Summary
Write-Host "`n✅ Members have been added to group:" -ForegroundColor Green
Write-Host "$($groupDisplayName)" -ForegroundColor Yellow
Write-Host "`n🎉 Finished processing all users!" -ForegroundColor Cyan
