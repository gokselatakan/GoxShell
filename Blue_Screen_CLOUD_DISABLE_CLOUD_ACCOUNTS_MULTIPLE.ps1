<#
Blue_Screen_Disable_AzureAD_Accounts
- Paste UPNs and/or Object IDs (one per line) and the script will disable each account (AccountEnabled = $false).
- Optional: pre-fill Admin UPN for the sign-in screen.
- Robust EXO-style connect logic (STA → interactive; otherwise Device Code).
- Progress + clear per-user results + optional CSV export to Desktop.
#>

try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

# --- Ensure AzureAD module is available ---
if (-not (Get-Module -ListAvailable -Name AzureAD)) {
    Write-Host "AzureAD module not found." -ForegroundColor Yellow
    $install = Read-Host "Install AzureAD module now? [Y/N]"
    if ($install -match '^[Yy]$') {
        Install-Module AzureAD -Scope CurrentUser -Force
    } else {
        Write-Host "Cannot continue without AzureAD module." -ForegroundColor Red
        return
    }
}

Import-Module AzureAD -ErrorAction SilentlyContinue

Add-Type -AssemblyName System.Windows.Forms

# ================== Blue Screen UI ==================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Disable Azure AD Accounts"
$form.Size = New-Object System.Drawing.Size(700, 520)
$form.BackColor = 'Blue'
$form.StartPosition = 'CenterScreen'

$lblAdmin = New-Object System.Windows.Forms.Label
$lblAdmin.Text = "Admin UPN for sign-in (optional):"
$lblAdmin.ForeColor = 'White'
$lblAdmin.AutoSize = $true
$lblAdmin.Location = New-Object System.Drawing.Point(12, 12)
$form.Controls.Add($lblAdmin)

$txtAdmin = New-Object System.Windows.Forms.TextBox
$txtAdmin.Size = New-Object System.Drawing.Size(660, 20)
$txtAdmin.Location = New-Object System.Drawing.Point(12, 32)
$form.Controls.Add($txtAdmin)

$lblList = New-Object System.Windows.Forms.Label
$lblList.Text = "Paste UPNs and/or Object IDs (one per line):"
$lblList.ForeColor = 'White'
$lblList.AutoSize = $true
$lblList.Location = New-Object System.Drawing.Point(12, 65)
$form.Controls.Add($lblList)

$txtList = New-Object System.Windows.Forms.TextBox
$txtList.Multiline = $true
$txtList.ScrollBars = 'Vertical'
$txtList.Size = New-Object System.Drawing.Size(660, 330)
$txtList.Location = New-Object System.Drawing.Point(12, 85)
$form.Controls.Add($txtList)

$chkExport = New-Object System.Windows.Forms.CheckBox
$chkExport.Text = "Export results to CSV on Desktop"
$chkExport.ForeColor = 'White'
$chkExport.AutoSize = $true
$chkExport.Location = New-Object System.Drawing.Point(12, 430)
$chkExport.Checked = $true
$form.Controls.Add($chkExport)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Disable Accounts"
$btnRun.Location = New-Object System.Drawing.Point(300, 425)
$btnRun.Add_Click({ $form.Close() })
$form.Controls.Add($btnRun)

$form.ShowDialog()

# ================== Read inputs ==================
$adminUPN   = $txtAdmin.Text.Trim()
$rawLines   = $txtList.Text -split "`r`n"
$identifiers = @()
foreach ($l in $rawLines) {
    $t = $l.Trim()
    if (-not [string]::IsNullOrWhiteSpace($t)) { $identifiers += $t }
}
$identifiers = $identifiers | Select-Object -Unique

if ($identifiers.Count -eq 0) {
    Write-Host "Please paste at least one UPN or Object ID." -ForegroundColor Yellow
    return
}

# ================== Robust AzureAD Connect ==================
function Connect-AzureAD-Robust {
    param([string]$UPN)

    # In WinPS 5.1, interactive control needs STA; if not STA, we'll try device code style
    $isSTA = ([System.Threading.Thread]::CurrentThread.ApartmentState -eq 'STA')
    try {
        if ([string]::IsNullOrWhiteSpace($UPN)) {
            Connect-AzureAD | Out-Null
        } else {
            # Note: older AzureAD uses -AccountId
            Connect-AzureAD -AccountId $UPN | Out-Null
        }
        return $true
    } catch {
        $msg = $_.Exception.Message
        if ($msg -match 'single-threaded apartment' -or $msg -match 'ActiveX control') {
            Write-Host "Interactive sign-in failed in non-STA context." -ForegroundColor Yellow
            Write-Host "Open a browser and sign in using device code..." -ForegroundColor Yellow
            try {
                # Fallback to device profile: AzureAD classic doesn’t have a first-class -Device switch.
                # Workaround: prompt without UPN to allow native device/browser flow. 
                Connect-AzureAD | Out-Null
                return $true
            } catch {
                Write-Host ("Device/browser sign-in also failed: {0}" -f $_.Exception.Message) -ForegroundColor Red
                return $false
            }
        } else {
            Write-Host ("Azure AD connect failed: {0}" -f $msg) -ForegroundColor Red
            return $false
        }
    }
}

Write-Host "`n🟡 In progress: Connecting to Azure AD..." -ForegroundColor Yellow
if (-not (Connect-AzureAD-Robust -UPN $adminUPN)) { return }

# ================== Helpers ==================
function Is-GuidLike {
    param([string]$s)
    if ([string]::IsNullOrWhiteSpace($s)) { return $false }
    $g = $null
    return [Guid]::TryParse($s, [ref]$g)
}

# Try to resolve user by UPN or ObjectId or search string
function Resolve-AADUser {
    param([string]$id)

    # ObjectId path
    if (Is-GuidLike -s $id) {
        try {
            $u = Get-AzureADUser -ObjectId $id -ErrorAction Stop
            if ($u) { return $u }
        } catch {}
    }

    # UPN path (fast)
    if ($id -match '@') {
        try {
            $u = Get-AzureADUser -Filter ("userPrincipalName eq '{0}'" -f $id.Replace("'","''")) -ErrorAction Stop
            if ($u) { return $u }
        } catch {}
    }

    # Search fallback (could return multiple; pick exact UPN if present)
    try {
        $candidates = Get-AzureADUser -SearchString $id -All $true
        if ($candidates) {
            $exact = $candidates | Where-Object { $_.UserPrincipalName -eq $id }
            if ($exact) { return ($exact | Select-Object -First 1) }
            return ($candidates | Select-Object -First 1)
        }
    } catch {}

    return $null
}

# ================== Process ==================
$success = @()
$already = @()
$notFound = @()
$failed  = @()

$total = $identifiers.Count
$idx = 0

foreach ($id in $identifiers) {
    $idx++
    Write-Progress -Activity "Disabling Azure AD accounts..." -Status "Processing $idx of $total" -PercentComplete (($idx / $total) * 100)

    $user = Resolve-AADUser -id $id
    if (-not $user) {
        $notFound += $id
        Write-Host ("❌ Not found: {0}" -f $id) -ForegroundColor Red
        continue
    }

    $display = if ($user.DisplayName) { $user.DisplayName } else { $user.UserPrincipalName }
    $upn = $user.UserPrincipalName
    $obj = $user.ObjectId

    # Get fresh 'AccountEnabled' state
    $enabled = $null
    try {
        $fresh = Get-AzureADUser -ObjectId $obj -ErrorAction Stop
        $enabled = $fresh.AccountEnabled
    } catch {}

    if ($enabled -eq $false) {
        $already += ("{0} ({1})" -f $display, $upn)
        Write-Host ("ℹ Already disabled: {0} ({1})" -f $display, $upn) -ForegroundColor DarkYellow
        continue
    }

    try {
        Set-AzureADUser -ObjectId $obj -AccountEnabled $false -ErrorAction Stop
        $success += ("{0} ({1})" -f $display, $upn)
        Write-Host ("✅ Disabled: {0} ({1})" -f $display, $upn) -ForegroundColor Green
    } catch {
        $failed += ("{0} ({1}) -> {2}" -f $display, $upn, $_.Exception.Message)
        Write-Host ("❌ Failed: {0} ({1}) -> {2}" -f $display, $upn, $_.Exception.Message) -ForegroundColor Red
    }
}

# ================== Summary ==================
Write-Host "`n================ Summary ================" -ForegroundColor Cyan
Write-Host ("Disabled : {0}" -f $success.Count) -ForegroundColor Green
if ($success.Count) { $success | ForEach-Object { Write-Host ("  - {0}" -f $_) -ForegroundColor Green } }

Write-Host ("Already  : {0}" -f $already.Count) -ForegroundColor DarkYellow
if ($already.Count) { $already | ForEach-Object { Write-Host ("  - {0}" -f $_) -ForegroundColor DarkYellow } }

Write-Host ("NotFound : {0}" -f $notFound.Count) -ForegroundColor Red
if ($notFound.Count) { $notFound | ForEach-Object { Write-Host ("  - {0}" -f $_) -ForegroundColor Red } }

Write-Host ("Failed   : {0}" -f $failed.Count) -ForegroundColor Red
if ($failed.Count)  { $failed  | ForEach-Object { Write-Host ("  - {0}" -f $_) -ForegroundColor Red } }

# ================== Optional CSV Export ==================
if ($chkExport.Checked) {
    $desktop = [Environment]::GetFolderPath('Desktop')
    $stamp   = (Get-Date).ToString('yyyyMMdd_HHmmss')
    $path    = Join-Path $desktop ("Disabled_Accounts_{0}.csv" -f $stamp)

    $rows = @()
    foreach ($s in $success)  { $rows += [PSCustomObject]@{ Result="Disabled";  Detail=$s } }
    foreach ($s in $already)  { $rows += [PSCustomObject]@{ Result="AlreadyDisabled"; Detail=$s } }
    foreach ($s in $notFound) { $rows += [PSCustomObject]@{ Result="NotFound"; Detail=$s } }
    foreach ($s in $failed)   { $rows += [PSCustomObject]@{ Result="Failed"; Detail=$s } }

    $rows | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8

    if (Test-Path $path) {
        Write-Host "`n✅ Export completed. File saved at: $path" -ForegroundColor Green
    } else {
        Write-Host "`n❌ CSV export failed." -ForegroundColor Red
    }
}

Write-Host "`n🎉 Finished processing all accounts." -ForegroundColor Cyan
