<#
Blue_Screen_AZURE_ADD_MEMBER_TO_MULTIPLE_GROUPS (User or Group)
- Enter a member identifier that can be:
    • User UPN (e.g., user@contoso.com), OR
    • Group Display Name (for nested group scenarios)
- Paste target Azure AD Group Display Names (one per line).
- Script auto-detects member type (User vs Group) and adds accordingly.
- Skips existing memberships; shows progress and summary.
#>

# --- STA bootstrap for WinForms ---
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    if (-not $PSCommandPath) {
        Write-Host "Please save this as a .ps1 and run again, or start PowerShell with: powershell.exe -STA" -ForegroundColor Yellow
        return
    }
    Write-Host "Restarting in STA mode..." -ForegroundColor Yellow
    $exe = (Get-Process -Id $PID).Path
    $argsList = @('-NoProfile','-ExecutionPolicy','Bypass','-STA','-File',"`"$PSCommandPath`"")
    Start-Process -FilePath $exe -ArgumentList $argsList -WindowStyle Normal | Out-Null
    exit
}

# --- Ensure AzureAD module ---
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
Import-Module AzureAD -Force

# --- UI ---
Add-Type -AssemblyName System.Windows.Forms

$form = New-Object System.Windows.Forms.Form
$form.Text = "Azure AD - Add Member (User or Group) to Multiple Groups"
$form.Size = New-Object System.Drawing.Size(600, 520)
$form.BackColor = 'Blue'
$form.StartPosition = 'CenterScreen'

$lblAdmin = New-Object System.Windows.Forms.Label
$lblAdmin.Text = "Admin UPN for sign-in (optional):"
$lblAdmin.ForeColor = 'White'
$lblAdmin.AutoSize = $true
$lblAdmin.Location = New-Object System.Drawing.Point(10, 12)
$form.Controls.Add($lblAdmin)

$txtAdmin = New-Object System.Windows.Forms.TextBox
$txtAdmin.Size = New-Object System.Drawing.Size(560, 20)
$txtAdmin.Location = New-Object System.Drawing.Point(10, 32)
$form.Controls.Add($txtAdmin)

$lblMember = New-Object System.Windows.Forms.Label
$lblMember.Text = "Member to add: User UPN (user@contoso.com) OR Group Display Name"
$lblMember.ForeColor = 'White'
$lblMember.AutoSize = $true
$lblMember.Location = New-Object System.Drawing.Point(10, 65)
$form.Controls.Add($lblMember)

$txtMember = New-Object System.Windows.Forms.TextBox
$txtMember.Size = New-Object System.Drawing.Size(560, 20)
$txtMember.Location = New-Object System.Drawing.Point(10, 85)
$form.Controls.Add($txtMember)

$lblGroups = New-Object System.Windows.Forms.Label
$lblGroups.Text = "Target Azure AD Group Display Names (one per line):"
$lblGroups.ForeColor = 'White'
$lblGroups.AutoSize = $true
$lblGroups.Location = New-Object System.Drawing.Point(10, 120)
$form.Controls.Add($lblGroups)

$txtGroups = New-Object System.Windows.Forms.TextBox
$txtGroups.Multiline = $true
$txtGroups.ScrollBars = 'Vertical'
$txtGroups.Size = New-Object System.Drawing.Size(560, 300)
$txtGroups.Location = New-Object System.Drawing.Point(10, 140)
$form.Controls.Add($txtGroups)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Add Member to Groups"
$btnRun.Location = New-Object System.Drawing.Point(230, 450)
$btnRun.Add_Click({ $form.Close() })
$form.Controls.Add($btnRun)

$form.ShowDialog()

# --- Inputs ---
$adminUPN    = $txtAdmin.Text.Trim()
$memberInput = $txtMember.Text.Trim()
$groupNames  = $txtGroups.Text -split "`r`n" | Where-Object { $_.Trim() -ne "" }

if ([string]::IsNullOrWhiteSpace($memberInput) -or $groupNames.Count -eq 0) {
    Write-Host "Please enter the member (UPN or Group Display Name) and at least one target group." -ForegroundColor Yellow
    return
}

# --- Connect to Azure AD ---
Write-Host "`n🔄 Connecting to Azure AD..." -ForegroundColor Cyan
if ([string]::IsNullOrWhiteSpace($adminUPN)) {
    Connect-AzureAD
} else {
    Connect-AzureAD -AccountId $adminUPN
}

# --- Helper: Resolve member (User or Group) ---
function Resolve-DirectoryObject {
    param([string]$identifier)

    # If it looks like an email, try user by UPN first (fastest/most exact)
    if ($identifier -match '@') {
        try {
            $u = Get-AzureADUser -Filter "userPrincipalName eq '$identifier'" -ErrorAction Stop
            if ($u) {
                return [PSCustomObject]@{
                    ObjectId   = $u.ObjectId
                    Type       = 'User'
                    Display    = $u.DisplayName
                    Canonical  = $u.UserPrincipalName
                }
            }
        } catch {}
    }

    # Try exact user match by SearchString (displayName/mail/etc.)
    try {
        $u2 = Get-AzureADUser -SearchString $identifier | Select-Object -First 1
        if ($u2) {
            return [PSCustomObject]@{
                ObjectId   = $u2.ObjectId
                Type       = 'User'
                Display    = $u2.DisplayName
                Canonical  = $u2.UserPrincipalName
            }
        }
    } catch {}

    # Try GROUP by exact displayName
    try {
        $g = Get-AzureADGroup -Filter "displayName eq '$identifier'"
        if ($g) {
            return [PSCustomObject]@{
                ObjectId   = $g.ObjectId
                Type       = 'Group'
                Display    = $g.DisplayName
                Canonical  = $g.DisplayName
            }
        }
    } catch {}

    # Try GROUP fuzzy search
    try {
        $g2 = Get-AzureADGroup -SearchString $identifier | Select-Object -First 1
        if ($g2) {
            return [PSCustomObject]@{
                ObjectId   = $g2.ObjectId
                Type       = 'Group'
                Display    = $g2.DisplayName
                Canonical  = $g2.DisplayName
            }
        }
    } catch {}

    return $null
}

$memberObj = Resolve-DirectoryObject -identifier $memberInput
if (-not $memberObj) {
    Write-Host "❌ Could not resolve member as User or Group: $memberInput" -ForegroundColor Red
    return
}

Write-Host "➡ Member resolved as $($memberObj.Type): $($memberObj.Canonical)" -ForegroundColor Cyan

# --- Process groups ---
$success = @()
$skipped = @()
$failed  = @()

$total = $groupNames.Count
$idx = 0

foreach ($gName in $groupNames) {
    $idx++
    Write-Progress -Activity "Adding member to groups..." -Status "Processing: $gName ($idx of $total)" -PercentComplete (($idx / $total) * 100)

    # Find target group by exact DisplayName first
    $group = $null
    try { $group = Get-AzureADGroup -Filter "displayName eq '$gName'" } catch {}
    if (-not $group) {
        try { $group = Get-AzureADGroup -SearchString $gName | Select-Object -First 1 } catch {}
    }
    if (-not $group) {
        $failed += "$gName (group not found)"
        Write-Host "❌ Group not found: $gName" -ForegroundColor Red
        continue
    }

    # Prevent self-nesting (group into itself)
    if ($memberObj.Type -eq 'Group' -and $memberObj.ObjectId -eq $group.ObjectId) {
        $skipped += "$gName (skipped: cannot add a group to itself)"
        Write-Host "ℹ Skipped (cannot add group to itself): $gName" -ForegroundColor DarkYellow
        continue
    }

    try {
        # Check existing membership (works for users or groups)
        $already = $false
        try {
            $existing = Get-AzureADGroupMember -ObjectId $group.ObjectId -All $true | Where-Object { $_.ObjectId -eq $memberObj.ObjectId }
            if ($existing) { $already = $true }
        } catch {}

        if ($already) {
            $skipped += "$gName (already a member)"
            Write-Host "ℹ Already a member of: $gName" -ForegroundColor DarkYellow
            continue
        }

        # Add member (user or group)
        Add-AzureADGroupMember -ObjectId $group.ObjectId -RefObjectId $memberObj.ObjectId
        $success += $gName

        if ($memberObj.Type -eq 'Group') {
            Write-Host "✅ Added GROUP '$($memberObj.Canonical)' to: $gName" -ForegroundColor Green
        } else {
            Write-Host "✅ Added USER '$($memberObj.Canonical)' to: $gName" -ForegroundColor Green
        }
    }
    catch {
        # Common cause: M365 (Unified) groups generally don't support nested groups
        $msg = $_.Exception.Message
        if ($memberObj.Type -eq 'Group' -and ($group.GroupTypes -and ($group.GroupTypes -contains 'Unified'))) {
            $msg = "$msg (note: Microsoft 365 groups typically do NOT allow nested groups)"
        }
        $failed += "$gName (error: $msg)"
        Write-Host "❌ Failed for ${gName}: $msg" -ForegroundColor Red
    }
}

# --- Summary ---
Write-Host "`n================ Summary ================" -ForegroundColor Cyan
Write-Host ("Added   : {0}" -f ($success.Count)) -ForegroundColor Green
if ($success.Count) { $success | ForEach-Object { Write-Host "  - $_" -ForegroundColor Green } }

Write-Host ("Skipped : {0}" -f ($skipped.Count)) -ForegroundColor DarkYellow
if ($skipped.Count) { $skipped | ForEach-Object { Write-Host "  - $_" -ForegroundColor DarkYellow } }

Write-Host ("Failed  : {0}" -f ($failed.Count)) -ForegroundColor Red
if ($failed.Count) { $failed  | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red } }

Write-Host "`n🎉 Finished processing all groups." -ForegroundColor Cyan
