Add-Type -AssemblyName System.Windows.Forms

# === First Form: Owner Display Name ===
$form1 = New-Object System.Windows.Forms.Form
$form1.Text = "Enter Owner Display Name"
$form1.Size = New-Object System.Drawing.Size(400, 150)
$form1.BackColor = 'Blue'

$label1 = New-Object System.Windows.Forms.Label
$label1.Text = "Enter the new owner's Display Name (e.g., John Doe):"
$label1.AutoSize = $true
$label1.Location = New-Object System.Drawing.Point(10, 10)
$form1.Controls.Add($label1)

$textBox1 = New-Object System.Windows.Forms.TextBox
$textBox1.Size = New-Object System.Drawing.Size(360, 20)
$textBox1.Location = New-Object System.Drawing.Point(10, 35)
$form1.Controls.Add($textBox1)

$okButton1 = New-Object System.Windows.Forms.Button
$okButton1.Text = "Next"
$okButton1.Location = New-Object System.Drawing.Point(150, 70)
$okButton1.Add_Click({ $form1.Close() })
$form1.Controls.Add($okButton1)

$form1.ShowDialog()
$ownerName = $textBox1.Text.Trim()

# === Second Form: Group List ===
$form2 = New-Object System.Windows.Forms.Form
$form2.Text = "Enter Group Names"
$form2.Size = New-Object System.Drawing.Size(400, 300)
$form2.BackColor = 'Blue'

$label2 = New-Object System.Windows.Forms.Label
$label2.Text = "Enter group names (one per line):"
$label2.AutoSize = $true
$label2.Location = New-Object System.Drawing.Point(10, 10)
$form2.Controls.Add($label2)

$textBox2 = New-Object System.Windows.Forms.TextBox
$textBox2.Multiline = $true
$textBox2.Size = New-Object System.Drawing.Size(360, 180)
$textBox2.Location = New-Object System.Drawing.Point(10, 35)
$textBox2.ScrollBars = 'Vertical'
$form2.Controls.Add($textBox2)

$okButton2 = New-Object System.Windows.Forms.Button
$okButton2.Text = "Update Owners"
$okButton2.Location = New-Object System.Drawing.Point(150, 230)
$okButton2.Add_Click({ $form2.Close() })
$form2.Controls.Add($okButton2)

$form2.ShowDialog()
$groupNames = $textBox2.Text -split "`r`n" | Where-Object { $_.Trim() -ne "" }

# === Load AD Module ===
Import-Module ActiveDirectory

# === Get Owner DN ===
$owner = Get-ADUser -Filter {DisplayName -eq $ownerName} -Properties DistinguishedName -ErrorAction SilentlyContinue
if (-not $owner) {
    Write-Host "❌ Owner not found: $ownerName" -ForegroundColor Red
    exit
}
$ownerDN = $owner.DistinguishedName

# === Update Groups ===
foreach ($groupName in $groupNames) {
    try {
        $group = Get-ADGroup -Filter {Name -eq $groupName} -ErrorAction Stop
        if ($group) {
            Set-ADGroup -Identity $group.DistinguishedName -ManagedBy $ownerDN
            Write-Host "✅ Updated: $($group.Name)" -ForegroundColor Green
        }
    } catch {
        Write-Host "❌ Error updating group `${groupName}`:`n$($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host "`n🎉 All done. Groups updated." -ForegroundColor Cyan
