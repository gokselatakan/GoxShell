<# 
Blue_Screen_Export_Leaver_Owned_Groups_To_CSV (EXO optional)
- On-Prem AD: ManagedBy = user  (Security / Distribution / Mail-enabled Security)
- Azure AD / M365: user is Owner (Security / M365 / Mail-enabled)
- Exchange Online Distribution Groups: user is Owner (OPTIONAL; can be slow)
- One CSV to Desktop
#>

# --- STA bootstrap (needed for WinForms/legacy prompts) ---
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    if (-not $PSCommandPath) {
        Write-Host "Save as .ps1 and run again, or start PowerShell with: powershell.exe -STA" -ForegroundColor Yellow
        return
    }
    Write-Host "Restarting in STA mode..." -ForegroundColor Yellow
    $exe = (Get-Process -Id $PID).Path
    $argsList = @('-NoProfile','-ExecutionPolicy','Bypass','-STA','-File',"`"$PSCommandPath`"")
    Start-Process -FilePath $exe -ArgumentList $argsList -WindowStyle Normal | Out-Null
    exit
}

# --- Prefer TLS 1.2 for EXO reliability ---
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

Add-Type -AssemblyName System.Windows.Forms

# ================== Blue Screen UI ==================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Export Leaver's Owned Groups - CSV"
$form.Size = New-Object System.Drawing.Size(620, 270)
$form.BackColor = 'Blue'
$form.StartPosition = 'CenterScreen'

$lblUser = New-Object System.Windows.Forms.Label
$lblUser.Text = "Enter leaver's sAMAccountName or UPN (email):"
$lblUser.ForeColor = 'White'
$lblUser.AutoSize = $true
$lblUser.Location = New-Object System.Drawing.Point(12, 12)
$form.Controls.Add($lblUser)

$txtUser = New-Object System.Windows.Forms.TextBox
$txtUser.Size = New-Object System.Drawing.Size(580, 20)
$txtUser.Location = New-Object System.Drawing.Point(12, 35)
$form.Controls.Add($txtUser)

$lblAdmin = New-Object System.Windows.Forms.Label
$lblAdmin.Text = "Cloud admin UPN (prefill Azure/EXO sign-in; optional):"
$lblAdmin.ForeColor = 'White'
$lblAdmin.AutoSize = $true
$lblAdmin.Location = New-Object System.Drawing.Point(12, 70)
$form.Controls.Add($lblAdmin)

$txtAdmin = New-Object System.Windows.Forms.TextBox
$txtAdmin.Size = New-Object System.Drawing.Size(580, 20)
$txtAdmin.Location = New-Object System.Drawing.Point(12, 93)
$form.Controls.Add($txtAdmin)

$chkExo = New-Object System.Windows.Forms.CheckBox
$chkExo.Text = "Include Exchange Online Distribution Groups (can be slow)"
$chkExo.AutoSize = $true
$chkExo.ForeColor = 'White'
$chkExo.Location = New-Object System.Drawing.Point(12, 125)
$chkExo.Checked = $false   # default OFF for speed
$form.Controls.Add($chkExo)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Export CSV"
$btnRun.Location = New-Object System.Drawing.Point(260, 170)
$btnRun.Add_Click({ $form.Close() })
$form.Controls.Add($btnRun)

$form.ShowDialog()

$userInput     = $txtUser.Text.Trim()
$AdminCloudUPN = $txtAdmin.Text.Trim()
$IncludeEXO    = $chkExo.Checked

if ([string]::IsNullOrWhiteSpace($userInput)) {
    Write-Host "⚠ Please enter a leaver user. Exiting." -ForegroundColor Yellow
    return
}

# ================== Output Prep ==================
$desktop = [Environment]::GetFolderPath('Desktop')
$stamp   = (Get-Date).ToString('yyyyMMdd_HHmmss')
$outFile = Join-Path $desktop ("Leaver_Owned_Groups_{0}_{1}.csv" -f ($userInput -replace '[\\/:*?""<>|]','_'), $stamp)

# ================== Dedup Store ==================
$items = @{}   # key: lower(Name) -> row
function Add-Row {
    param([string]$Name,[string]$Type,[string]$Source,[string]$Description,[string[]]$Owners)
    if ([string]::IsNullOrWhiteSpace($Name)) { return }
    $key = $Name.ToLowerInvariant()
    $ownerStr = ($Owners | Where-Object { $_ -and $_.Trim() } | Select-Object -Unique) -join '; '
    if ($items.ContainsKey($key)) {
        $cur = $items[$key]
        $srcs = ($cur.'Source(s)'.ToString() -split ';').ForEach({ $_.Trim() }) | Where-Object { $_ }
        if ($srcs -notcontains $Source) { $cur.'Source(s)' = ($srcs + $Source) -join '; ' }
        if ([string]::IsNullOrWhiteSpace($cur.Description) -and $Description) { $cur.Description = $Description }
        $own = ($cur.'Owner(s)'.ToString() -split ';').ForEach({ $_.Trim() }) | Where-Object { $_ }
        $cur.'Owner(s)' = (( $own + ($ownerStr -split ';') ) | Where-Object { $_ } | Select-Object -Unique) -join '; '
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

# ================== Modules ==================
$adAvailable = $true
try { Import-Module ActiveDirectory -ErrorAction Stop } catch { $adAvailable = $false; Write-Host "❌ RSAT ActiveDirectory not found; on-prem will be skipped." -ForegroundColor Red }

$aadAvailable = $true
if (-not (Get-Module -ListAvailable -Name AzureAD)) {
    $aadAvailable = $false
    Write-Host "AzureAD module not found; cloud ownership via AzureAD will be skipped." -ForegroundColor Yellow
} else { Import-Module AzureAD -ErrorAction SilentlyContinue }

$exoAvailable = $IncludeEXO -and (Get-Module -ListAvailable -Name ExchangeOnlineManagement)
if (-not $exoAvailable -and $IncludeEXO) {
    Write-Host "ExchangeOnlineManagement module not found; EXO ownership will be skipped." -ForegroundColor Yellow
}

# ================== ON-PREM: groups where user is ManagedBy ==================
if ($adAvailable) {
    Write-Host "`n🟡 In progress: Collecting ON-PREM groups the user OWNS (ManagedBy)..." -ForegroundColor Yellow
    $adUser = $null
    try { $adUser = Get-ADUser -Identity $userInput -Properties DistinguishedName,DisplayName -ErrorAction Stop } catch {}
    if (-not $adUser) {
        try { $adUser = Get-ADUser -Filter "UserPrincipalName -eq '$userInput'" -Properties DistinguishedName,DisplayName -ErrorAction Stop } catch {}
    }

    if ($adUser -and $adUser.DistinguishedName) {
        $ownerDN = $adUser.DistinguishedName
        $ownerName = if ($adUser.DisplayName) { $adUser.DisplayName } else { $adUser.Name }
        try {
            $ownedGroups = Get-ADGroup -Filter { ManagedBy -eq $ownerDN } -Properties Name, GroupCategory, Mail, Description
            $total = ($ownedGroups | Measure-Object).Count
            $i = 0
            foreach ($g in $ownedGroups) {
                $i++; $pct = if ($total) { [int](($i/$total)*100) } else { 100 }
                Write-Progress -Activity "On-Prem AD (Owned)" -Status "Processing $i of $total" -PercentComplete $pct
                $type = if ($g.GroupCategory -eq 'Distribution') { 'Distribution Group (On-Prem)' }
                        elseif ($g.Mail) { 'Mail-enabled Security (On-Prem)' }
                        else { 'Security Group (On-Prem)' }
                Add-Row -Name $g.Name -Type $type -Source 'On-Prem AD' -Description $g.Description -Owners @($ownerName)
            }
        } catch {
            Write-Host ("❌ Error querying on-prem groups: {0}" -f $_.Exception.Message) -ForegroundColor Red
        }
    } else {
        Write-Host "ℹ On-prem user not found; skipping on-prem ownership." -ForegroundColor DarkYellow
    }
}

# ================== AZURE AD: groups where user is Owner ==================
if ($aadAvailable) {
    Write-Host "`n🟡 In progress: Connecting to Azure AD..." -ForegroundColor Yellow
    $aadConnected = $false
    try {
        if ([string]::IsNullOrWhiteSpace($AdminCloudUPN)) { Connect-AzureAD | Out-Null } else { Connect-AzureAD -AccountId $AdminCloudUPN | Out-Null }
        $aadConnected = $true
    } catch { Write-Host ("❌ Azure AD connect failed: {0}" -f $_.Exception.Message) -ForegroundColor Red }

    if ($aadConnected) {
        $cloudUser = $null
        try { $cloudUser = Get-AzureADUser -Filter "userPrincipalName eq '$userInput'" -ErrorAction Stop } catch {}
        if (-not $cloudUser) {
            try { $cloudUser = Get-AzureADUser -SearchString $userInput | Select-Object -First 1 } catch {}
        }

        if ($cloudUser) {
            Write-Host "🟡 In progress: Collecting CLOUD groups the user OWNS..." -ForegroundColor Yellow
            try {
                $owned = Get-AzureADUserOwnedObject -ObjectId $cloudUser.ObjectId -All $true | Where-Object { $_.ObjectType -eq 'Group' }
                $totalA = ($owned | Measure-Object).Count
                $j = 0
                foreach ($g in $owned) {
                    $j++; $pct = if ($totalA) { [int](($j/$totalA)*100) } else { 100 }
                    Write-Progress -Activity "Azure AD (Owned)" -Status "Processing $j of $totalA" -PercentComplete $pct

                    $ag = $null
                    try { $ag = Get-AzureADGroup -ObjectId $g.ObjectId } catch {}
                    if ($ag) {
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
                        if (-not $owners -or $owners.Count -eq 0) { $owners = @($cloudUser.DisplayName) }

                        Add-Row -Name $ag.DisplayName -Type $type -Source 'Azure AD / Cloud' -Description $ag.Description -Owners $owners
                    }
                }
            } catch {
                Write-Host ("❌ Error while reading Azure AD owned groups: {0}" -f $_.Exception.Message) -ForegroundColor Red
            }
        } else {
            Write-Host "ℹ User not found in Azure AD; skipping cloud ownership." -ForegroundColor DarkYellow
        }
    }
}

# ================== EXCHANGE ONLINE (OPTIONAL): Distribution Groups ownership ==================
if ($IncludeEXO -and (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "`n🟡 In progress: Connecting to Exchange Online..." -ForegroundColor Yellow
    $exoConnected = $false
    try {
        if ([string]::IsNullOrWhiteSpace($AdminCloudUPN)) {
            Connect-ExchangeOnline -ShowBanner:$false | Out-Null
        } else {
            Connect-ExchangeOnline -UserPrincipalName $AdminCloudUPN -ShowBanner:$false | Out-Null
        }
        $exoConnected = $true
    } catch { Write-Host ("❌ EXO connect failed: {0}" -f $_.Exception.Message) -ForegroundColor Red }

    if ($exoConnected) {
        Write-Host "🟡 In progress: Collecting EXO Distribution Groups the user OWNS (this may take time)..." -ForegroundColor Yellow
        $exoUser = $null
        try { $exoUser = Get-User -Identity $userInput -ErrorAction Stop } catch {}
        $ownerDN = if ($exoUser) { $exoUser.DistinguishedName } else { $null }

        try {
            $dgs = Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop
            $totalDG = ($dgs | Measure-Object).Count
            $k = 0
            foreach ($dg in $dgs) {
                $k++; $pct = if ($totalDG) { [int](($k/$totalDG)*100) } else { 100 }
                if (($k % 50) -eq 0) { Write-Progress -Activity "EXO (Owned DGs)" -Status "Scanning $k of $totalDG" -PercentComplete $pct }

                $isOwner = $false
                $owners = @()

                if ($dg.ManagedBy -and $dg.ManagedBy.Count -gt 0) {
                    foreach ($mgr in $dg.ManagedBy) {
                        if ($ownerDN -and ($mgr -eq $ownerDN)) { $isOwner = $true }
                        try {
                            $mgrObj = Get-User -Identity $mgr -ErrorAction SilentlyContinue
                            if ($mgrObj -and $mgrObj.DisplayName) { $owners += $mgrObj.DisplayName }
                            elseif ($mgrObj -and $mgrObj.Name) { $owners += $mgrObj.Name }
                        } catch {}
                    }
                }

                if ($isOwner) {
                    Add-Row -Name $dg.DisplayName -Type "Distribution Group (Cloud/EXO)" -Source "Exchange Online" -Description $dg.Notes -Owners $owners
                }
            }
        } catch {
            Write-Host ("❌ Error while reading EXO distribution groups: {0}" -f $_.Exception.Message) -ForegroundColor Red
        }

        try { Disconnect-ExchangeOnline -Confirm:$false | Out-Null } catch {}
    }
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
