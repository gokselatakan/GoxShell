<#
Blue_Screen_Bulk_SoftDelete_AzureAD_Accounts (Fixed GUID check)
- Paste UPNs and/or Object IDs (one per line).
- Optional admin UPN prefill for the sign-in screen.
- Deletes with Remove-AzureADUser (soft delete; remains recoverable ~30 days).
- Progress + per-item status + optional CSV summary to Desktop.
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
$form.Text = "Bulk Soft Delete Azure AD Accounts (Cloud)"
$form.Size = New-Object System.Drawing.Size(760, 560)
$form.BackColor = 'Blue'
$form.StartPosition = 'CenterScreen'

# Admin UPN prefill
$lblAdmin = New-Object System.Windows.Forms.Label
$lblAdmin.Text = "Admin UPN for sign-in (optional):"
$lblAdmin.ForeColor = 'White'
$lblAdmin.AutoSize = $true
$lblAdmin.Location = New-Object System.Drawing.Point(12, 12)
$form.Controls.Add($lblAdmin)

$txtAdmin = New-Object System.Windows.Forms.TextBox
$txtAdmin.Size = New-Object System.Drawing.Size(720, 20)
$txtAdmin.Location = New-Object System.Drawing.Point(12, 32)
$form.Controls.Add($txtAdmin)

# List label + textbox
$lblList = New-Object System.Windows.Forms.Label
$lblList.Text = "Paste UPNs and/or Object IDs (one per line):"
$lblList.ForeColor = 'White'
$lblList.AutoSize = $true
$lblList.Location = New-Object System.Drawing.Point(12, 65)
$form.Controls.Add($lblList)

$txtList = New-Object System.Windows.Forms.TextBox
$txtList.Multiline = $true
$txtList.ScrollBars = 'Vertical'
$txtList.Size = New-Object System.Drawing.Size(720, 360)
$txtList.Location = New-Object System.Drawing.Point(12, 85)
$form.Controls.Add($txtList)

$chkExport = New-Object System.Windows.Forms.CheckBox
$chkExport.Text = "Export results to CSV on Desktop"
$chkExport.ForeColor = 'White'
$chkExport.AutoSize = $true
$chkExport.Location = New-Object System.Drawing.Point(12, 460)
$chkExport.Checked = $true
$form.Controls.Add($chkExport)

# Run button
$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "SOFT DELETE Accounts"
$btnRun.Location = New-Object System.Drawing.Point(320, 455)
$btnRun.Add_Click({ $form.Close() })
$form.Controls.Add($btnRun)

$form.ShowDialog()

# ================== Read inputs ==================
$adminUPN    = $txtAdmin.Text.Trim()
$rawLines    = $txtList.Text -split "`r`n"
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

# ================== Connect (robust) ==================
function Connect-AzureAD-Robust {
    param([string]$UPN)
    try {
        if ([string]::IsNullOrWhiteSpace($UPN)) {
            Connect-AzureAD | Out-Null
        } else {
            Connect-AzureAD -AccountId $UPN | Out-Null
        }
        return $true
    } catch {
        $msg = $_.Exception.Message
        if ($msg -match 'single-threaded apartment' -or $msg -match 'ActiveX control') {
            Write-Host "Interactive sign-in failed in non-STA context." -ForegroundColor Yellow
            Write-Host "Open a browser and sign in (device/browser flow)..." -ForegroundColor Yellow
            try {
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
    [Guid]$tmp = [Guid]::Empty
    return [Guid]::TryParse($s, [ref]$tmp)
}

function Resolve-AADUser {
    param([string]$id)

    if (Is-GuidLike -s $id) {
        try { $u = Get-AzureADUser -ObjectId $id -ErrorAction Stop; if ($u) { return $u } } catch {}
    }

    if ($id -match '@') {
        try {
            $safe = $id.Replace("'","''")
            $u = Get-AzureADUser -Filter ("userPrincipalName eq '{0}'" -f $safe) -ErrorAction Stop
            if ($u) { return $u }
        } catch {}
    }

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
$deleted  = @()
$notFound = @()
$failed   = @()

$total = $identifiers.Count
$idx = 0

foreach ($id in $identifiers) {
    $idx++
    Write-Progress -Activity "Soft-deleting Azure AD accounts..." -Status ("Processing {0} of {1}" -f $idx, $total) -PercentComplete (($idx / $total) * 100)

    $user = Resolve-AADUser -id $id
    if (-not $user) {
        $notFound += $id
        Write-Host ("❌ Not found: {0}" -f $id) -ForegroundColor Red
        continue
    }

    $upn     = $user.UserPrincipalName
    $objId   = $user.ObjectId
    $display = if ($user.DisplayName) { $user.DisplayName } else { $upn }

    try {
        Remove-AzureADUser -ObjectId $objId -ErrorAction Stop   # soft delete
        $deleted += ("{0} ({1})" -f $display, $upn)
        Write-Host ("✅ Soft-deleted: {0} ({1})" -f $display, $upn) -ForegroundColor Green
    } catch {
        $failed += ("{0} ({1}) -> {2}" -f $display, $upn, $_.Exception.Message)
        Write-Host ("❌ Failed: {0} ({1}) -> {2}" -f $display, $upn, $_.Exception.Message) -ForegroundColor Red
    }
}

# ================== Summary ==================
Write-Host "`n================ Summary ================" -ForegroundColor Cyan
Write-Host ("Soft-deleted : {0}" -f $deleted.Count) -ForegroundColor Green
if ($deleted.Count) { $deleted | ForEach-Object { Write-Host ("  - {0}" -f $_) -ForegroundColor Green } }

Write-Host ("Not found    : {0}" -f $notFound.Count) -ForegroundColor Red
if ($notFound.Count) { $notFound | ForEach-Object { Write-Host ("  - {0}" -f $_) -ForegroundColor Red } }

Write-Host ("Failed       : {0}" -f $failed.Count) -ForegroundColor Red
if ($failed.Count)  { $failed  | ForEach-Object { Write-Host ("  - {0}" -f $_) -ForegroundColor Red } }

# ================== Optional CSV Export ==================
if ($chkExport.Checked) {
    $desktop = [Environment]::GetFolderPath('Desktop')
    $stamp   = (Get-Date).ToString('yyyyMMdd_HHmmss')
    $path    = Join-Path $desktop ("SoftDeleted_Accounts_{0}.csv" -f $stamp)

    $rows = @()
    foreach ($s in $deleted)  { $rows += [PSCustomObject]@{ Result="SoftDeleted";  Detail=$s } }
    foreach ($s in $notFound) { $rows += [PSCustomObject]@{ Result="NotFound";    Detail=$s } }
    foreach ($s in $failed)   { $rows += [PSCustomObject]@{ Result="Failed";      Detail=$s } }

    $rows | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8

    if (Test-Path $path) {
        Write-Host "`n✅ Export completed. File saved at: $path" -ForegroundColor Green
    } else {
        Write-Host "`n❌ CSV export failed." -ForegroundColor Red
    }
}

Write-Host "`n🎉 Finished processing soft deletions. (Users remain in 'Deleted users' for ~30 days.)" -ForegroundColor Cyan
