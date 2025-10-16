# Load Windows Forms
Add-Type -AssemblyName System.Windows.Forms

# Create Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Get Email Addresses from Display Names"
$form.Size = New-Object System.Drawing.Size(500, 400)
$form.BackColor = 'Blue'

# Instructions Label
$label = New-Object System.Windows.Forms.Label
$label.Text = "Paste display names below (one per line):"
$label.ForeColor = 'White'
$label.Location = New-Object System.Drawing.Point(10, 10)
$label.AutoSize = $true
$form.Controls.Add($label)

# Multiline TextBox for input
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Multiline = $true
$textBox.ScrollBars = 'Vertical'
$textBox.Size = New-Object System.Drawing.Size(460, 250)
$textBox.Location = New-Object System.Drawing.Point(10, 35)
$form.Controls.Add($textBox)

# OK Button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "Get Email Addresses"
$okButton.Location = New-Object System.Drawing.Point(180, 300)
$okButton.Add_Click({ $form.Close() })
$form.Controls.Add($okButton)

# Show the Form
$form.ShowDialog()

# Get Display Names from Textbox
$displayNames = $textBox.Text -split "`r`n" | Where-Object { $_.Trim() -ne "" }

# Load AD module
Import-Module ActiveDirectory

# Prepare results array
$results = @()

# Process each display name
foreach ($name in $displayNames) {
    $user = Get-ADUser -Filter "DisplayName -eq '$name'" -Properties Mail -ErrorAction SilentlyContinue
    if ($user) {
        $results += [PSCustomObject]@{
            DisplayName = $name
            Email       = $user.Mail
        }
    } else {
        $results += [PSCustomObject]@{
            DisplayName = $name
            Email       = "❌ Not Found"
        }
    }
}

# Export results to CSV
$outputPath = "C:\Users\$env:USERNAME\Desktop\DisplayName_Email_Results.csv"
$results | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8

# Confirmation
Write-Host "`n✅ Email lookup completed." -ForegroundColor Green
Write-Host "📄 Results saved to: $outputPath" -ForegroundColor Cyan
