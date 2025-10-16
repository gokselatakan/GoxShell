<# 
Blue_Screen_Export_Mover_All_Access_CSV_Only (Resilient EXO)

Features
- Blue screens prompt for:
    • Target user's sAMAccountName or UPN (mover)
    • Cloud admin UPN to prefill Microsoft sign-in (you'll enter the password)
- On-Prem AD groups (ManagedBy -> Display Name, Description)
- Azure AD / M365 groups (Owners -> Display Names, Description)
- Exchange Online Shared Mailboxes:
    • FullAccess / SendAs (server-side filtered)
    • SendOnBehalf (only if shared mailbox directory loads; otherwise skipped)
- Deduplicates by Name; merges Source(s) and Owner(s)
- Exports one CSV to Desktop; shows in-progress + green "Exported" message
#>

# Ensure TLS 1.2 for EXO REST
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
} catch {}

Add-Type -AssemblyName System.Windows.Forms

# ================== Blue Screen UI ==================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Export All Access (On-Prem + Cloud + Shared Mailboxes) - CSV"
$form.Size = New-Object System.Drawing.Size(560, 210)
$form.BackColor = 'Blue'
$form.StartPosition = 'CenterScreen'

$lblUser = New-Object System.Windows.Forms.Label
$lblUser.Text = "Enter mover user's sAMAccountName or UPN:"
$lblUser.ForeColor = 'White'
$lblUser.AutoSize = $true
$lblUser.Location = New-Object System.Drawing.Point(12, 12)
$form.Controls.Add($lblUser)

$txtUser = New-Object System.Windows.Forms.TextBox
$txtUser.Size = New-Object System.Drawing.Size(520, 20)
$txtUser.Location = New-Object System.Drawing.Point(12, 35)
$form.Controls.Add($txtUser)

$lblAdmin = New-Object System.Windows.Forms.Label
$lblAdmin.Text = "Cloud admin UPN (prefill Azure/Exchange sign-in):"
$lblAdmin.ForeColor = 'White'
$lblAdmin.AutoSize = $true
$lblAdmin.Location = New-Object System.Drawing.Point(12, 70)
$form.Controls.Add($lblAdmin)

$txtAdmin = New-Object System.Windows.Forms.TextBox
$txtAdmin.Size = New-Object System.Drawing.Size(520, 20)
$txtAdmin.Location = New-Object System.Drawing.Point(12, 93)
$form.Controls.Add($txtAdmin)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Export CSV"
$btnRun.Location = New-Object System.Drawing.Point(235, 130)
$btnRun.Add_Click({ $form.Close() })
$form.Controls.Add($btnRun)

$form.ShowDialog()

$userInput     = $txtUser.Text.Trim()
$AdminCloudUPN = $txtAdmin.Text.Trim()

if ([string]::IsNullOrWhiteSpace($userInput)) {
    Write-Host "⚠ Please enter a mover user. Exiting." -ForegroundColor Yellow
    return
}

# ================== Output Prep ==================
$desktop = [Environment]::GetFolderPath('Desktop')
$stamp   = (Get-Date).ToString('yyyyMMdd_HHmmss')
$outFile = Join-Path $desktop ("Mover_Access_{0}_{1}.csv" -f ($userInput -replace '[\\/:*?""<>|]','_'), $stamp)

# Master map keyed by lowercase Name to de-duplicate across sources
$items = @{}   # key: nameLower -> PSCustomObject

function Add-Row {
    param(
        [string]$Name,
        [string]$Type,
        [string]$Source,
        [string]$Description,
        [string[]]$Owners
    )
    if ([string]::IsNullOrWhiteSpace($Name)) { return }
    $key = $Name.ToLowerInvariant()
    $ownerStr = ($Owners | Where-Object { $_ -and $_.Trim() -ne "" } | Select-Object -Unique) -join '; '
    if ($items.ContainsKey($key)) {
        $cur = $items[$key]
        # merge sources
        $srcs = ($cur.'Source(s)'.ToString() -split ';').ForEach({ $_.Trim() }) | Where-Object { $_ }
        if ($srcs -notcontains $Source) { $cur.'Source(s)' = ($srcs + $Source) -join '; ' }
        # richer description wins
        if ([string]::IsNullOrWhiteSpace($cur.Description) -and -not [string]::IsNullOrWhiteSpace($Description)) {
            $cur.Description = $Description
        }
        # merge owners
        $own = ($cur.'Owner(s)'.ToString() -split ';').ForEach({ $_.Trim() }) | Where-Object { $_ }
        $mergedOwners = ( ($own + ($ownerStr -split ';')).ForEach({ $_.Trim() }) | Where-Object { $_ } | Select-Object -Unique ) -join '; '
        $cur.'Owner(s)' = $mergedOwners
        # keep original Type unless empty
        if ([string]::IsNullOrWhiteSpace($cur.'Type')) { $cur.'Type' = $Type }
        $items[$key] = $cur
    } else {
        $items[$key] = [PSCustomObject]@{
            'Name'        = $Name
            'Type'        = $Type
            'Source(s)'   = $Source
            'Description' = $Description
            'Owner(s)'    = $ownerStr
        }
    }
}

# ================== Module Checks ==================
# On-Prem AD
try { Import-Module ActiveDirectory -ErrorAction Stop } catch {
    Write-Host "❌ ActiveDirectory module not found (RSAT). On-prem portion will be skipped." -ForegroundColor Red
}

# AzureAD
if (-not (Get-Module -ListAvailable -Name AzureAD)) {
    Write-Host "AzureAD module not found." -ForegroundColor Yellow
    $ans = Read-Host "Install AzureAD module now? [Y/N]"
    if ($ans -match '^[Yy]$') { Install-Module AzureAD -Scope CurrentUser -Force }
}
try { Import-Module AzureAD -Force } catch {}

# Exchange Online
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "ExchangeOnlineManagement module not found." -ForegroundColor Yellow
    $ans2 = Read-Host "Install Exchange Online module now? [Y/N]"
    if ($ans2 -match '^[Yy]$') { Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force }
}

# ================== ON-PREM: groups ==================
if (Get-Module -Name ActiveDirectory) {
    Write-Host "`n🟡 In progress: Collecting ON-PREM AD groups..." -ForegroundColor Yellow

    $adUser = $null
    try { $adUser = Get-ADUser -Identity $userInput -Properties MemberOf -ErrorAction Stop } catch {}
    if (-not $adUser) {
        try { $adUser = Get-ADUser -Filter "UserPrincipalName -eq '$userInput'" -Properties MemberOf -ErrorAction Stop } catch {}
    }

    if ($adUser -and $adUser.MemberOf) {
        $total = $adUser.MemberOf.Count
        $i = 0
        foreach ($dn in $adUser.MemberOf) {
            $i++
            $pct = [int](($i / $total) * 100)
            Write-Progress -Activity "On-Prem AD" -Status "Resolving group $i of $total" -PercentComplete $pct
            try {
                $g = Get-ADGroup -Identity $dn -Properties Name, GroupCategory, Mail, ManagedBy, Description -ErrorAction Stop
                $type = if ($g.GroupCategory -eq 'Distribution') { 'Distribution Group (On-Prem)' }
                        elseif ($g.Mail) { 'Mail-enabled Security (On-Prem)' }
                        else { 'Security Group (On-Prem)' }

                $owners = @()
                if ($g.ManagedBy) {
                    try {
                        $ownU = Get-ADUser -Identity $g.ManagedBy -Properties DisplayName -ErrorAction Stop
                        if ($ownU.DisplayName) { $owners += $ownU.DisplayName } else { $owners += $ownU.Name }
                    } catch {
                        try {
                            $ownObj = Get-ADObject -Identity $g.ManagedBy -Properties Name -ErrorAction Stop
                            $owners += $ownObj.Name
                        } catch {}
                    }
                }

                Add-Row -Name $g.Name -Type $type -Source 'On-Prem AD' -Description $g.Description -Owners $owners
            } catch {}
        }
    } else {
        Write-Host "ℹ On-prem user not found or no direct groups." -ForegroundColor DarkYellow
    }
}

# ================== AZURE AD: groups ==================
if (Get-Module -Name AzureAD) {
    Write-Host "`n🟡 In progress: Connecting to Azure AD..." -ForegroundColor Yellow
    try {
        if ([string]::IsNullOrWhiteSpace($AdminCloudUPN)) { Connect-AzureAD }
        else { Connect-AzureAD -AccountId $AdminCloudUPN }
    } catch {
        Write-Host "❌ Azure AD connect failed: $($_.Exception.Message)" -ForegroundColor Red
    }

    $cloudUser = $null
    try { $cloudUser = Get-AzureADUser -Filter "userPrincipalName eq '$userInput'" -ErrorAction Stop } catch {}
    if (-not $cloudUser) { try { $cloudUser = Get-AzureADUser -SearchString $userInput | Select-Object -First 1 } catch {} }

    if ($cloudUser) {
        Write-Host "🟡 In progress: Collecting AZURE AD / M365 groups..." -ForegroundColor Yellow
        try {
            $mems = Get-AzureADUserMembership -ObjectId $cloudUser.ObjectId -All $true | Where-Object { $_.ObjectType -eq 'Group' }
            $totalA = ($mems | Measure-Object).Count
            $j = 0
            foreach ($m in $mems) {
                $j++
                $percent = if ($totalA -gt 0) { [int](($j / $totalA) * 100) } else { 100 }
                Write-Progress -Activity "Azure AD" -Status "Resolving group $j of $totalA" -PercentComplete $percent

                $ag = $null
                try { $ag = Get-AzureADGroup -ObjectId $m.ObjectId } catch {}
                if ($ag) {
                    $gName = $ag.DisplayName
                    $isUnified = ($ag.GroupTypes -and ($ag.GroupTypes -contains 'Unified'))
                    $type =
                        if ($isUnified) { 'Microsoft 365 Group (Cloud)' }
                        elseif ($ag.MailEnabled -and -not $ag.SecurityEnabled) { 'Distribution Group (Cloud)' }
                        elseif ($ag.MailEnabled -and $ag.SecurityEnabled) { 'Mail-enabled Security (Cloud)' }
                        else { 'Security Group (Cloud)' }

                    $owners = @()
                    try {
                        $ownObjs = Get-AzureADGroupOwner -ObjectId $ag.ObjectId -All $true
                        foreach ($o in $ownObjs) {
                            if ($o.DisplayName) { $owners += $o.DisplayName }
                            elseif ($o.UserPrincipalName) { $owners += $o.UserPrincipalName }
                            elseif ($o.Mail) { $owners += $o.Mail }
                        }
                    } catch {}

                    Add-Row -Name $gName -Type $type -Source 'Azure AD / Cloud' -Description $ag.Description -Owners $owners
                }
            }
        } catch {
            Write-Host "❌ Error while reading Azure AD memberships: $($_.Exception.Message)" -ForegroundColor Red
        }
    } else {
        Write-Host "ℹ Cloud user not found in Azure AD (skipping cloud groups)." -ForegroundColor DarkYellow
    }
}

# ================== EXCHANGE ONLINE: shared mailboxes (fast path + retry + fallback) ==================
if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
    Write-Host "`n🟡 In progress: Connecting to Exchange Online..." -ForegroundColor Yellow
    try {
        if ([string]::IsNullOrWhiteSpace($AdminCloudUPN)) { Connect-ExchangeOnline -ShowBanner:$false }
        else { Connect-ExchangeOnline -UserPrincipalName $AdminCloudUPN -ShowBanner:$false }
    } catch {
        Write-Host "❌ EXO connect failed: $($_.Exception.Message)" -ForegroundColor Red
    }

    $userForPerm = $userInput
    $sharedMbx = $null

    function Get-SharedMailboxDirectory {
        param([int]$MaxAttempts = 3, [int]$DelaySeconds = 2)
        $attempt = 0
        do {
            $attempt++
            try {
                # Prefer EXO REST
                return Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -PropertySets All -ErrorAction Stop
            } catch {
                if ($attempt -lt $MaxAttempts) {
                    Start-Sleep -Seconds $DelaySeconds
                } else {
                    throw
                }
            }
        } while ($attempt -lt $MaxAttempts)
    }

    try {
        Write-Host "🟡 In progress: Loading shared mailbox directory (retry-enabled)..." -ForegroundColor Yellow
        $sharedMbx = Get-SharedMailboxDirectory
    } catch {
        Write-Host "⚠ Could not load shared mailbox directory via Get-EXOMailbox. Trying fallback..." -ForegroundColor DarkYellow
        try {
            # Fallback: legacy cmdlet (also REST-backed in newer modules but tends to work when Get-EXO… hiccups)
            $sharedMbx = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -ErrorAction Stop
        } catch {
            Write-Host "❌ Fallback Get-Mailbox also failed. SendOnBehalf scanning will be skipped." -ForegroundColor Red
            $sharedMbx = $null
        }
    }

    # Build lookups if directory loaded
    $smbxIndexBySmtp = @{}
    $smbxIndexByName = @{}
    if ($sharedMbx) {
        foreach ($m in $sharedMbx) {
            if ($m.PrimarySmtpAddress) { $smbxIndexBySmtp[$m.PrimarySmtpAddress.ToLower()] = $m }
            if ($m.DisplayName)        { $smbxIndexByName[$m.DisplayName.ToLower()]        = $m }
        }
    }

    try {
        # ------------ FullAccess (server-side filtered) ------------
        Write-Host "🟡 In progress: Resolving FullAccess (server-side filtered)..." -ForegroundColor Yellow
        $faPerms = @()
        try {
            $faPerms = Get-EXOMailboxPermission -User $userForPerm -ResultSize Unlimited -ErrorAction Stop |
                       Where-Object {
                           $_.AccessRights -contains 'FullAccess' -and
                           -not $_.IsInherited -and
                           $_.User -ne 'NT AUTHORITY\SELF' -and
                           $_.User -ne 'S-1-5-10'
                       }
        } catch {}

        foreach ($p in $faPerms) {
            $mbx = $null
            $id = "$($p.Identity)".ToLower()
            if ($smbxIndexBySmtp.ContainsKey($id))      { $mbx = $smbxIndexBySmtp[$id] }
            elseif ($smbxIndexByName.ContainsKey($id))  { $mbx = $smbxIndexByName[$id] }
            else {
                try {
                    $tmp = Get-EXOMailbox -Identity $p.Identity -ErrorAction Stop
                    if ($tmp.RecipientTypeDetails -eq 'SharedMailbox') { $mbx = $tmp }
                } catch {
                    try {
                        $tmp = Get-Mailbox -Identity $p.Identity -ErrorAction Stop
                        if ($tmp.RecipientTypeDetails -eq 'SharedMailbox') { $mbx = $tmp }
                    } catch {}
                }
            }
            if ($mbx) {
                $owners = @()
                try {
                    $u = Get-User -Identity $mbx.Identity
                    if ($u.Manager) {
                        try {
                            $mgr = Get-User -Identity $u.Manager
                            if ($mgr.DisplayName) { $owners += $mgr.DisplayName } else { $owners += $mgr.Name }
                        } catch {}
                    }
                } catch {}
                Add-Row -Name $mbx.DisplayName -Type "Shared Mailbox (FullAccess)" -Source "Exchange Online" -Description "" -Owners $owners
            }
        }

        # ------------ SendAs (server-side filtered) ------------
        Write-Host "🟡 In progress: Resolving SendAs (server-side filtered)..." -ForegroundColor Yellow
        $saPerms = @()
        try {
            $saPerms = Get-RecipientPermission -Trustee $userForPerm -ResultSize Unlimited -ErrorAction Stop |
                       Where-Object { -not $_.Deny }
        } catch {}

        foreach ($rp in $saPerms) {
            $mbx = $null
            $key = "$($rp.Identity)".ToLower()
            if ($smbxIndexBySmtp.ContainsKey($key))      { $mbx = $smbxIndexBySmtp[$key] }
            elseif ($smbxIndexByName.ContainsKey($key))  { $mbx = $smbxIndexByName[$key] }
            else {
                try {
                    $tmp = Get-EXOMailbox -Identity $rp.Identity -ErrorAction Stop
                    if ($tmp.RecipientTypeDetails -eq 'SharedMailbox') { $mbx = $tmp }
                } catch {
                    try {
                        $tmp = Get-Mailbox -Identity $rp.Identity -ErrorAction Stop
                        if ($tmp.RecipientTypeDetails -eq 'SharedMailbox') { $mbx = $tmp }
                    } catch {}
                }
            }
            if ($mbx) {
                $owners = @()
                try {
                    $u = Get-User -Identity $mbx.Identity
                    if ($u.Manager) {
                        try {
                            $mgr = Get-User -Identity $u.Manager
                            if ($mgr.DisplayName) { $owners += $mgr.DisplayName } else { $owners += $mgr.Name }
                        } catch {}
                    }
                } catch {}
                Add-Row -Name $mbx.DisplayName -Type "Shared Mailbox (SendAs)" -Source "Exchange Online" -Description "" -Owners $owners
            }
        }

        # ------------ SendOnBehalf (property filter if directory loaded) ------------
        if ($sharedMbx) {
            Write-Host "🟡 In progress: Resolving SendOnBehalf (property filter)..." -ForegroundColor Yellow
            foreach ($mb in $sharedMbx) {
                try {
                    if ($mb.GrantSendOnBehalfTo -and ($mb.GrantSendOnBehalfTo -contains $userForPerm)) {
                        $owners = @()
                        try {
                            $u = Get-User -Identity $mb.Identity
                            if ($u.Manager) {
                                try {
                                    $mgr = Get-User -Identity $u.Manager
                                    if ($mgr.DisplayName) { $owners += $mgr.DisplayName } else { $owners += $mgr.Name }
                                } catch {}
                            }
                        } catch {}
                        Add-Row -Name $mb.DisplayName -Type "Shared Mailbox (SendOnBehalf)" -Source "Exchange Online" -Description "" -Owners $owners
                    }
                } catch {}
            }
        } else {
            Write-Host "ℹ Skipping SendOnBehalf scan (shared mailbox directory unavailable)." -ForegroundColor DarkYellow
        }

    } catch {
        Write-Host "❌ Error collecting shared mailbox permissions: $($_.Exception.Message)" -ForegroundColor Red
    }

    try { Disconnect-ExchangeOnline -Confirm:$false | Out-Null } catch {}
}

# ================== Export CSV ==================
$result = $items.Values | Sort-Object Name
$result | Export-Csv -Path $outFile -NoTypeInformation -Encoding UTF8

if (Test-Path $outFile) {
    Write-Host "`n✅ Exported $($result.Count) items." -ForegroundColor Green
    Write-Host "📄 File saved at: $outFile" -ForegroundColor Cyan
} else {
    Write-Host "`n❌ Export failed." -ForegroundColor Red
}
