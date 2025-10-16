# ================================================
# Purpose:
#   Revoke all refresh tokens for a user with Microsoft Graph.
#   Prefer Revoke-MgUserSignInSession; if missing, POST /users/{id}/revokeSignInSessions
#   with an empty JSON body {} (required by Invoke-MgGraphRequest for POST).
#
# Flow:
#   1) Blue input window -> enter ObjectId or UPN
#   2) Read signInSessionsValidFromDateTime (before)
#   3) Revoke (cmdlet if present; else REST with {} body)
#   4) Read timestamp (after) and print green Success / red Failed
# ================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Blue input window ---
$form               = New-Object System.Windows.Forms.Form
$form.Text          = "Microsoft Graph - Revoke Refresh Tokens"
$form.Size          = New-Object System.Drawing.Size(480,240)
$form.StartPosition = "CenterScreen"
$form.BackColor     = [System.Drawing.Color]::MidnightBlue
$form.ForeColor     = [System.Drawing.Color]::White
$form.TopMost       = $true

$label = New-Object System.Windows.Forms.Label
$label.Text = "Enter the User ObjectId or UPN:"
$label.Location = New-Object System.Drawing.Point(16,20)
$label.AutoSize = $true
$label.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(19,50)
$textBox.Width = 430
$form.Controls.Add($textBox)

$runButton = New-Object System.Windows.Forms.Button
$runButton.Text = "Run"
$runButton.Width = 90
$runButton.Location = New-Object System.Drawing.Point(140,110)
$runButton.Add_Click({ $form.Tag = "RUN"; $form.Close() })
$form.Controls.Add($runButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = "Cancel"
$cancelButton.Width = 90
$cancelButton.Location = New-Object System.Drawing.Point(250,110)
$cancelButton.Add_Click({ $form.Tag = "CANCEL"; $form.Close() })
$form.Controls.Add($cancelButton)

$form.ShowDialog() | Out-Null
if ($form.Tag -eq "CANCEL") { Write-Host "Operation cancelled." -ForegroundColor Yellow; return }

$userId = $textBox.Text.Trim()
if (-not $userId) { Write-Host "❌ No ObjectId/UPN entered." -ForegroundColor Red; exit 1 }

try {
    # Try to import Users.Actions (contains Revoke-MgUserSignInSession)
    $null = Import-Module Microsoft.Graph.Users.Actions -ErrorAction SilentlyContinue
    Import-Module Microsoft.Graph.Users -ErrorAction SilentlyContinue | Out-Null

    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes "User.ReadWrite.All" -ErrorAction Stop

    # Read 'before' timestamp
    Write-Host "Fetching current signInSessionsValidFromDateTime..." -ForegroundColor Cyan
    $before = (Get-MgUser -UserId $userId -Property signInSessionsValidFromDateTime).signInSessionsValidFromDateTime

    # Revoke using cmdlet if available; else REST with {} body
    if (Get-Command -Name Revoke-MgUserSignInSession -ErrorAction SilentlyContinue) {
        Write-Host "Revoking via Revoke-MgUserSignInSession..." -ForegroundColor Cyan
        Revoke-MgUserSignInSession -UserId $userId -ErrorAction Stop
    } else {
        Write-Host "Cmdlet not found. Revoking via REST POST /users/{id}/revokeSignInSessions..." -ForegroundColor Cyan
        $uri = "https://graph.microsoft.com/v1.0/users/$([uri]::EscapeDataString($userId))/revokeSignInSessions"

        # Detect SDK variant: -BodyParameter (v2+) or -Body (older)
        $invokeParams = (Get-Command Invoke-MgGraphRequest).Parameters
        if ($invokeParams.ContainsKey('BodyParameter')) {
            Invoke-MgGraphRequest -Method POST -Uri $uri -BodyParameter @{} -ErrorAction Stop | Out-Null
        } else {
            Invoke-MgGraphRequest -Method POST -Uri $uri -Body '{}' -ContentType 'application/json' -ErrorAction Stop | Out-Null
        }
    }

    # Give Graph a moment to update
    Start-Sleep -Seconds 3

    # Read 'after' timestamp
    Write-Host "Fetching updated signInSessionsValidFromDateTime..." -ForegroundColor Cyan
    $after = (Get-MgUser -UserId $userId -Property signInSessionsValidFromDateTime).signInSessionsValidFromDateTime

    $beforeDt = if ($before) { [datetime]$before } else { $null }
    $afterDt  = if ($after)  { [datetime]$after }  else { $null }

    if ($afterDt -and (-not $beforeDt -or $afterDt -gt $beforeDt)) {
        Write-Host "✅ Success: refresh tokens were revoked." -ForegroundColor Green
        Write-Host ("Before : {0}" -f $beforeDt) -ForegroundColor Yellow
        Write-Host ("After  : {0}" -f $afterDt)  -ForegroundColor Yellow
    } else {
        Write-Host "❌ Failed or timestamp did not change." -ForegroundColor Red
        Write-Host ("Before : {0}" -f $beforeDt) -ForegroundColor Yellow
        Write-Host ("After  : {0}" -f $afterDt)  -ForegroundColor Yellow
    }
}
catch {
    Write-Host "❌ Error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
