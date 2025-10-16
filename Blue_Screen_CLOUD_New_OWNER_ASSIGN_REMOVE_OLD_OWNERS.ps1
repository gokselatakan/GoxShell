# Install module if needed
# Install-Module AzureAD -Scope CurrentUser -Force

# Load AzureAD Module
Import-Module AzureAD -Force

# Add GUI components
Add-Type -AssemblyName System.Windows.Forms

# Create Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Azure AD Group - Assign New Owner"
$form.Size = New-Object System.Drawing.Size(500, 400)
$form.BackColor = 'Blue'

# New Owner Email Label & TextBox
$labelOwner = New-Object System.Windows.Forms.Label
$labelOwner.Text = "Enter New Owner's Email Address:"
$labelOwner.ForeColor = 'White'
$labelOwner.Location = New-Object System.Drawing.Point(10, 10)
$labelOwner.AutoSize = $true
$form.Controls.Add($labelOwner)

$textOwner = New-Object System.Windows.Forms.TextBox
$textOwner.Size = New-Object System.Drawing.Size(460, 20)
$textOwner.Location = New-Object System.Drawing.Point(10, 30)
$form.Controls.Add($textOwner)

# Group List Label & Multiline TextBox
$labelGroups = New-Object System.Windows.Forms.Label
$labelGroups.Text = "Enter Azure AD Group Display Names (one per line):"
$labelGroups.ForeColor = 'White'
$labelGroups.Location = New-Object System.Drawing.Point(10, 65)
$labelGroups.AutoSize = $true
$form.Controls.Add($labelGroups)

$textGroups = New-Object System.Windows.Forms.TextBox
$textGroups.Multiline = $true
$textGroups.ScrollBars = 'Vertical'
$textGroups.Size = New-Object System.Drawing.Size(460, 200)
$textGroups.Location = New-Object System.Drawing.Point(10, 85)
$form.Controls.Add($textGroups)

# OK Button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "Assign New Owner"
$okButton.Location = New-Object System.Drawing.Point(180, 300)
$okButton.Add_Click({ $form.Close() })
$form.Controls.Add($okButton)

# Show Form
$form.ShowDialog()

# Read Inputs
$newOwnerEmail = $textOwner.Text.Trim()
$groupList = $textGroups.Text -split "`r`n" | Where-Object { $_.Trim() -ne "" }

# --- Connect to Azure AD ---
Write-Host "`n🔄 Connecting to Azure AD..." -ForegroundColor Cyan
if ([string]::IsNullOrWhiteSpace($adminUPN)) {
    Connect-AzureAD
} else {
    Connect-AzureAD -AccountId $adminUPN
}

# Get new owner user object
$newOwner = Get-AzureADUser -Filter "UserPrincipalName eq '$newOwnerEmail'" -ErrorAction SilentlyContinue

if (-not $newOwner) {
    Write-Host "❌ New owner not found: $newOwnerEmail" -ForegroundColor Red
    exit
}

# Process each group
foreach ($groupName in $groupList) {
    $group = Get-AzureADGroup -Filter "DisplayName eq '$groupName'" -ErrorAction SilentlyContinue

    if (-not $group) {
        Write-Host "❌ Group not found: $groupName" -ForegroundColor Red
        continue
    }

    try {
        # Assign new owner
        Add-AzureADGroupOwner -ObjectId $group.ObjectId -RefObjectId $newOwner.ObjectId
        Write-Host "✅ Assigned $newOwnerEmail as new owner to: $groupName" -ForegroundColor Green

        # Get existing owners
        $existingOwners = Get-AzureADGroupOwner -ObjectId $group.ObjectId | Where-Object { $_.ObjectId -ne $newOwner.ObjectId }

        foreach ($oldOwner in $existingOwners) {
            Remove-AzureADGroupOwner -ObjectId $group.ObjectId -OwnerId $oldOwner.ObjectId
            Write-Host "🗑️ Removed previous owner: $($oldOwner.UserPrincipalName)" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "❌ Error updating group ${groupName}: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host "`n🎉 Finished updating group ownership!" -ForegroundColor Cyan
