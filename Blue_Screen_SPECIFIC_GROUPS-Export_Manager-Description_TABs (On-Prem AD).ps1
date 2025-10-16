Add-Type -AssemblyName System.Windows.Forms

# Create Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Enter Security Groups"
$form.Size = New-Object System.Drawing.Size(400,300)
$form.BackColor = 'Blue'

# Create TextBox
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Multiline = $true
$textBox.Size = New-Object System.Drawing.Size(360,180)
$textBox.Location = New-Object System.Drawing.Point(10,10)
$textBox.ScrollBars = 'Vertical'
$form.Controls.Add($textBox)

# Create OK Button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Location = New-Object System.Drawing.Point(150,200)
$okButton.Add_Click({ $form.Close() })
$form.Controls.Add($okButton)

# Show Form
$form.ShowDialog()

# Get input groups
$groups = $textBox.Text -split "`r`n" | Where-Object { $_ -ne "" }

# Import Active Directory module
Import-Module ActiveDirectory

# Initialize array to store results
$results = @()

foreach ($group in $groups) {
    $adGroup = Get-ADGroup -Filter {Name -eq $group} -Properties ManagedBy, Description -ErrorAction SilentlyContinue
    if ($adGroup) {
        $manager = $adGroup.ManagedBy
        $description = $adGroup.Description
        
        # Extract common name (CN) only from distinguished name (DN)
        $cn = ($manager -split ",")[0] -replace "CN=", ""
        
        # Store results
        $results += [PSCustomObject]@{
            "Group Name" = $group
            "Manager" = $cn
            "Description" = $description
        }
    }
}

# Export to CSV instead of Excel due to security restrictions
$csvPath = "C:\Users\$env:USERNAME\Desktop\AD_Groups_Report.csv"
$results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

Write-Host "Report generated at: $csvPath" -ForegroundColor Green
