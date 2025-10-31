param(
    [ValidateSet("Prompt","HTML","CSV","Both")]
    [string]$ExportFormat = "Prompt",

    # İstersen sadece hiç method kaydetmemiş kullanıcıları dahil et
    [switch]$OnlyNoAuthMethods
)

# ----------------------------
# Connect to Microsoft Graph
# ----------------------------
Connect-MgGraph -Scopes "UserAuthenticationMethod.Read.All","User.Read.All" | Out-Null

# ----------------------------
# Get all users (+ createdDateTime)
# ----------------------------
$uri = "beta/users?`$select=Id,DisplayName,UserPrincipalName,createdDateTime"
$Result   = Invoke-MgGraphRequest -Uri $Uri -OutputType PSObject
$AllUsers = [System.Collections.Generic.List[Object]]::new()
$AllUsers.AddRange($Result.value)
$NextLink = $Result."@odata.nextLink"
while ($NextLink) {
    $Result = Invoke-MgGraphRequest -Method GET -Uri $NextLink -OutputType PSObject
    $AllUsers.AddRange($Result.value)
    $NextLink = $Result."@odata.nextLink"
}

# ----------------------------
# Helpers
# ----------------------------
function Get-LastUsedString {
    param($methods,$odataType)
    $dt = ($methods | Where-Object { $_."@odata.type" -eq $odataType }).lastUsedDateTime
    if ([string]::IsNullOrWhiteSpace($dt)) { return "" }
    try { return ([DateTime]$dt).ToString("yyyy-MM-dd HH:mm") } catch { return "" }
}

function Has-Method {
    param($methods,$odataType)
    return [bool]($methods | Where-Object { $_."@odata.type" -eq $odataType })
}

# ----------------------------
# Build data
# ----------------------------
$authMethodsReport = [System.Collections.Generic.List[Object]]::new()

foreach ($user in $AllUsers) {
    $methods = @()
    try {
        $muri    = "beta/users/$($user.Id)/authentication/methods"
        $methods = (Invoke-MgGraphRequest -Uri $muri -OutputType PSObject).value
    } catch { $methods = @() }

    $hasAny = ($methods -and $methods.Count -gt 0)
    if ($OnlyNoAuthMethods -and $hasAny) { continue }

    # Presence flags
    $hasFido2        = Has-Method $methods "#microsoft.graph.fido2AuthenticationMethod"
    $hasAuth         = Has-Method $methods "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod"
    $hasAuthPwdless  = Has-Method $methods "#microsoft.graph.passwordlessMicrosoftAuthenticatorAuthenticationMethod"
    $hasPhone        = Has-Method $methods "#microsoft.graph.phoneAuthenticationMethod"
    $hasSWOath       = Has-Method $methods "#microsoft.graph.softwareOathAuthenticationMethod"
    $hasWHfB         = Has-Method $methods "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod"

    # Last-used strings
    $luEmail         = Get-LastUsedString $methods "#microsoft.graph.emailAuthenticationMethod"
    $luExternal      = Get-LastUsedString $methods "#microsoft.graph.externalAuthenticationMethod"
    $luFido2         = Get-LastUsedString $methods "#microsoft.graph.fido2AuthenticationMethod"
    $luPassword      = Get-LastUsedString $methods "#microsoft.graph.passwordAuthenticationMethod"
    $luAuth          = Get-LastUsedString $methods "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod"
    $luAuthPwdless   = Get-LastUsedString $methods "#microsoft.graph.passwordlessMicrosoftAuthenticatorAuthenticationMethod"
    $luHWoath        = Get-LastUsedString $methods "#microsoft.graph.hardwareOathAuthenticationMethod"
    $luPhone         = Get-LastUsedString $methods "#microsoft.graph.phoneAuthenticationMethod"
    $luSWoath        = Get-LastUsedString $methods "#microsoft.graph.softwareOathAuthenticationMethod"
    $luTAP           = Get-LastUsedString $methods "#microsoft.graph.temporaryAccessPassAuthenticationMethod"
    $luWHfB          = Get-LastUsedString $methods "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod"
    $luPlatform      = Get-LastUsedString $methods "#microsoft.graph.platformCredentialAuthenticationMethod"
    $luQrPin         = Get-LastUsedString $methods "#microsoft.graph.qrCodePinAuthenticationMethod"

    # Never-used değerlendirmesi (var ama hiç kullanılmadı)
    $authNeverUsed = ($hasAuth -and [string]::IsNullOrWhiteSpace($luAuth))

    # Safety: FIDO2 var + Authenticator Never => Auth silme riskli
    $doNotRemoveAuthenticator   = ($hasFido2 -and $authNeverUsed)
    $safeToRemoveAuthenticator  = ($authNeverUsed -and -not $doNotRemoveAuthenticator)

    # MFA health flags
    $hasMfa = ($hasAuth -or $hasPhone -or $hasSWOath -or $hasAuthPwdless -or $hasFido2 -or $hasWHfB)
    $hasStrongMfa = ($hasFido2 -or $hasWHfB -or $hasAuthPwdless)

    # Created date
    $createdDateString = ""
    if ($user.createdDateTime) {
        try { $createdDateString = ([DateTime]$user.createdDateTime).ToString("yyyy-MM-dd HH:mm") } catch { $createdDateString = "" }
    }

    $obj = [PSCustomObject]@{
        UserPrincipalName             = $user.UserPrincipalName
        CreatedDate                   = $createdDateString

        HasAnyMethod                  = $hasAny
        MethodsCount                  = ($methods | Measure-Object).Count

        HasFido2Registered            = $hasFido2
        AuthenticatorRegistered       = $hasAuth
        HasAuthPasswordlessRegistered = $hasAuthPwdless
        HasPhoneRegistered            = $hasPhone
        HasSoftwareOathRegistered     = $hasSWOath
        HasWHfBRegistered             = $hasWHfB

        AuthenticatorNeverUsed        = $authNeverUsed
        DoNotRemoveAuthenticator      = $doNotRemoveAuthenticator
        SafeToRemoveAuthenticator     = $safeToRemoveAuthenticator

        HasMfa                        = $hasMfa
        HasStrongMfa                  = $hasStrongMfa

        Email                         = $luEmail
        External                      = $luExternal
        Fido2                         = $luFido2
        Password                      = $luPassword
        Authenticator                 = $luAuth
        AuthenticatorPasswordless     = $luAuthPwdless
        HardwareOath                  = $luHWoath
        Phone                         = $luPhone
        SoftwareOath                  = $luSWoath
        TemporaryAccessPass           = $luTAP
        WindowsHello                  = $luWHfB
        PlatformCredential            = $luPlatform
        QRCodePIN                     = $luQrPin
    }
    $authMethodsReport.Add($obj)
}

# ----------------------------
# MFA Health Score (summary)
# ----------------------------
$totalUsers     = $authMethodsReport.Count
$mfaUsers       = ($authMethodsReport | Where-Object { $_.HasMfa }).Count
$strongUsers    = ($authMethodsReport | Where-Object { $_.HasStrongMfa }).Count
$mfaPct         = if ($totalUsers -gt 0) { [math]::Round(($mfaUsers*100.0)/$totalUsers, 1) } else { 0 }
$strongPct      = if ($totalUsers -gt 0) { [math]::Round(($strongUsers*100.0)/$totalUsers, 1) } else { 0 }

# ----------------------------
# HTML (NO STICKY)
# ----------------------------
$reportDate = Get-Date -Format "MMMM dd, yyyy"
$reportTime = Get-Date -Format "HH:mm:ss"

$htmlHeader = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Authentication Methods Usage Report</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,'Helvetica Neue',Arial,sans-serif;line-height:1.4;color:#1a1a1a;background-color:#f8f9fa;padding:16px}
.container{max-width:100%;margin:0 auto;background-color:#fff;box-shadow:0 1px 3px rgba(0,0,0,.08);border-radius:4px;overflow:hidden}
.header{background-color:#2d3748;color:#fff;padding:20px 24px;border-bottom:1px solid #1a202c}
.header h1{font-size:22px;font-weight:600;margin-bottom:4px}
.header .meta{font-size:13px;color:#cbd5e0}
.controls{padding:12px 24px;background-color:#f7fafc;border-bottom:1px solid #e2e8f0;display:flex;flex-wrap:wrap;gap:16px;align-items:center}
.filter-group{display:flex;align-items:center;gap:8px}
.controls label{display:inline-flex;align-items:center;font-size:13px;color:#4a5568;cursor:pointer;margin:0}
.controls input[type="checkbox"]{margin-right:6px;cursor:pointer}
.controls input[type="text"],.controls select,.controls button{padding:6px 10px;border:1px solid #cbd5e0;border-radius:4px;font-size:13px;color:#2d3748;background-color:#fff}
.controls button{cursor:pointer}
.summary{padding:12px 24px;background-color:#ffffff;border-bottom:1px solid #e2e8f0;display:flex;flex-wrap:wrap;gap:24px;align-items:center}
.kpi{background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:10px 14px;min-width:260px}
.kpi .title{font-size:12px;color:#4a5568;margin-bottom:6px}
.kpi .value{font-weight:700;font-size:18px;color:#1a202c}
.progress{height:8px;background:#e5e7eb;border-radius:6px;margin-top:8px;overflow:hidden}
.progress > div{height:8px;background:#3b82f6;width:0%}
.kpi.small{min-width:200px}
.table-container{padding:0;overflow-x:auto}
table{width:100%;border-collapse:collapse;font-size:13px}
thead{background-color:#edf2f7}
th{padding:10px 8px;text-align:left;font-weight:600;color:#2d3748;font-size:12px;border-bottom:2px solid #cbd5e0;white-space:nowrap}
tbody tr{border-bottom:1px solid #e2e8f0}
tbody tr:hover{background-color:#f7fafc}
td{padding:8px;color:#2d3748;white-space:nowrap}
td:first-child{font-weight:500;color:#1a202c;position:sticky;left:0;background-color:#fff;max-width:320px;overflow:hidden;text-overflow:ellipsis}
.date-cell{font-family:'Consolas','Monaco',monospace;font-size:12px;color:#4a5568}
.never-used{color:#a0aec0;font-style:italic;font-size:12px}
.registered-never{color:#374151;font-style:italic;font-size:12px}
.badge{display:inline-block;padding:2px 6px;border-radius:8px;font-size:11px;margin-left:6px;border:1px solid #CBD5E0;color:#4A5568}
.badge.warn{border-color:#F59E0B;color:#9A3412;background:#FEF3C7}
.hidden-row{display:none!important}
.footer{padding:12px 24px;background-color:#f7fafc;border-top:1px solid #e2e8f0;text-align:center;font-size:12px;color:#718096}
.hidden-column{display:none}
@media print{.controls,.summary{display:none}}
</style>
</head>
<body>
<div class="container">
    <div class="header">
        <h1>Authentication Methods Usage Report</h1>
        <div class="meta">Generated on $reportDate at $reportTime</div>
    </div>

    <div class="controls">
        <div class="filter-group">
            <label><input type="checkbox" id="hideEmptyColumns" checked onchange="toggleEmptyColumns()">Hide empty method columns</label>
        </div>

        <div class="filter-group">
            <span class="filter-label">Search:</span>
            <input type="text" id="searchInput" placeholder="Search users..." onkeyup="applyFilters()">
        </div>

        <div class="filter-group">
            <span class="filter-label">Inactive for:</span>
            <select id="inactivityFilter" onchange="applyFilters()">
                <option value="">All</option>
                <option value="30">30+ days</option>
                <option value="60">60+ days</option>
                <option value="90">90+ days</option>
                <option value="180">180+ days</option>
                <option value="365">1+ year</option>
                <option value="never">Never used any method</option>
            </select>
        </div>

        <div class="filter-group">
            <span class="filter-label">Method:</span>
            <select id="methodFilter" onchange="applyFilters()">
                <option value="">All</option>
                <option value="Email">Email</option>
                <option value="External">External</option>
                <option value="Fido2">FIDO2</option>
                <option value="Password">Password</option>
                <option value="Authenticator">Authenticator</option>
                <option value="AuthenticatorPasswordless">Auth Passwordless</option>
                <option value="HardwareOath">Hardware OATH</option>
                <option value="Phone">Phone</option>
                <option value="SoftwareOath">Software OATH</option>
                <option value="TemporaryAccessPass">Temp Access Pass</option>
                <option value="WindowsHello">Windows Hello</option>
                <option value="PlatformCredential">Platform Credential</option>
                <option value="QRCodePIN">QR Code PIN</option>
            </select>
            <select id="methodStatusFilter" onchange="applyFilters()">
                <option value="">Any status</option>
                <option value="has">Has been used</option>
                <option value="never">Never used</option>
            </select>
        </div>

        <div class="filter-group">
            <label><input type="checkbox" id="protectFido2Auth" checked onchange="applyFilters()">Protect when FIDO2 present & Authenticator shows Never</label>
        </div>

        <div class="filter-group">
            <span class="filter-label">Sort by:</span>
            <select id="sortField" onchange="applyFilters()">
                <option value="">None</option>
                <option value="CreatedDate">Created Date</option>
                <option value="MostRecent">Most recent activity (any method)</option>
                <option value="Authenticator">Authenticator last used</option>
                <option value="Fido2">FIDO2 last used</option>
            </select>
            <select id="sortDirection" onchange="applyFilters()">
                <option value="desc">Desc</option>
                <option value="asc">Asc</option>
            </select>
            <label><input type="checkbox" id="neverFirst" onchange="applyFilters()">Place Never first (by Created Date)</label>
        </div>

        <div class="filter-group">
            <button id="exportCsvBtn" onclick="exportVisibleToCsv()">Export visible to Excel (CSV)</button>
        </div>
    </div>

    <div class="summary">
        <div class="kpi">
            <div class="title">MFA Enrolled Users</div>
            <div class="value">$mfaUsers / $totalUsers ($mfaPct%)</div>
            <div class="progress"><div style="width:$mfaPct%"></div></div>
            <div style="font-size:12px;color:#6b7280;margin-top:6px">Counted when any of: Authenticator, Phone, Software OATH, Passwordless, FIDO2, Windows Hello</div>
        </div>
        <div class="kpi">
            <div class="title">Strong MFA (FIDO2 / WHfB / Auth Passwordless)</div>
            <div class="value">$strongUsers / $totalUsers ($strongPct%)</div>
            <div class="progress"><div style="width:$strongPct%"></div></div>
            <div style="font-size:12px;color:#6b7280;margin-top:6px">Phishing-resistant methods share</div>
        </div>
        <div class="kpi small">
            <div class="title">Report date/time</div>
            <div class="value">$reportDate $reportTime</div>
        </div>
    </div>

    <div class="table-container">
        <table id="dataTable">
            <thead>
                <tr>
                    <th data-column="0">User Principal Name</th>
                    <th data-column="Z">Created Date</th>

                    <th data-column="A">HasAnyMethod</th>
                    <th data-column="B">MethodsCount</th>

                    <th data-column="C">HasFido2Registered</th>
                    <th data-column="D">AuthenticatorRegistered</th>
                    <th data-column="H">HasAuthPasswordlessRegistered</th>
                    <th data-column="P">HasPhoneRegistered</th>
                    <th data-column="S">HasSoftwareOathRegistered</th>
                    <th data-column="W">HasWHfBRegistered</th>

                    <th data-column="E">AuthenticatorNeverUsed</th>
                    <th data-column="F">DoNotRemoveAuthenticator</th>
                    <th data-column="G">SafeToRemoveAuthenticator</th>

                    <th data-column="MFA">HasMfa</th>
                    <th data-column="STRONG">HasStrongMfa</th>

                    <th data-column="1">Email</th>
                    <th data-column="2">External</th>
                    <th data-column="3">FIDO2</th>
                    <th data-column="4">Password</th>
                    <th data-column="5">Authenticator</th>
                    <th data-column="6">Auth Passwordless</th>
                    <th data-column="7">Hardware OATH</th>
                    <th data-column="8">Phone</th>
                    <th data-column="9">Software OATH</th>
                    <th data-column="10">Temp Access Pass</th>
                    <th data-column="11">Windows Hello</th>
                    <th data-column="12">Platform Credential</th>
                    <th data-column="13">QR Code PIN</th>
                </tr>
            </thead>
            <tbody>
"@

# ----------------------------
# HTML BODY
# ----------------------------
$htmlBody = ""
foreach ($user in $authMethodsReport) {
    $badge = ""
    if ($user.DoNotRemoveAuthenticator) { $badge = "<span class='badge warn'>FIDO2 present — keep Authenticator</span>" }
    elseif ($user.SafeToRemoveAuthenticator) { $badge = "<span class='badge'>Safe to remove Authenticator</span>" }

    $htmlBody += "                <tr>`n"
    $htmlBody += "                    <td data-column='0'>$($user.UserPrincipalName) $badge</td>`n"
    $htmlBody += "                    <td data-column='Z'>$($user.CreatedDate)</td>`n"

    $htmlBody += "                    <td data-column='A'>$($user.HasAnyMethod)</td>`n"
    $htmlBody += "                    <td data-column='B'>$($user.MethodsCount)</td>`n"

    $htmlBody += "                    <td data-column='C'>$($user.HasFido2Registered)</td>`n"
    $htmlBody += "                    <td data-column='D'>$($user.AuthenticatorRegistered)</td>`n"
    $htmlBody += "                    <td data-column='H'>$($user.HasAuthPasswordlessRegistered)</td>`n"
    $htmlBody += "                    <td data-column='P'>$($user.HasPhoneRegistered)</td>`n"
    $htmlBody += "                    <td data-column='S'>$($user.HasSoftwareOathRegistered)</td>`n"
    $htmlBody += "                    <td data-column='W'>$($user.HasWHfBRegistered)</td>`n"

    $htmlBody += "                    <td data-column='E'>$($user.AuthenticatorNeverUsed)</td>`n"
    $htmlBody += "                    <td data-column='F'>$($user.DoNotRemoveAuthenticator)</td>`n"
    $htmlBody += "                    <td data-column='G'>$($user.SafeToRemoveAuthenticator)</td>`n"

    $htmlBody += "                    <td data-column='MFA'>$($user.HasMfa)</td>`n"
    $htmlBody += "                    <td data-column='STRONG'>$($user.HasStrongMfa)</td>`n"

    $properties = @('Email','External','Fido2','Password','Authenticator','AuthenticatorPasswordless',
                    'HardwareOath','Phone','SoftwareOath','TemporaryAccessPass','WindowsHello',
                    'PlatformCredential','QRCodePIN')
    $colIndex = 1
    foreach ($prop in $properties) {
        $value = $user.$prop
        $dataStatus = "never"
        $class = "never-used"
        $text = "Never"

        if (-not [string]::IsNullOrWhiteSpace($value)) {
            $dataStatus = "used"
            $class = "date-cell"
            $text = $value
        } else {
            if ($prop -eq 'Authenticator' -and ($user.AuthenticatorRegistered -or $user.HasAuthPasswordlessRegistered)) {
                $dataStatus = "registered-never"
                $class = "registered-never"
                $text = "Registered (never used)"
            }
        }

        $htmlBody += "                    <td class='$class' data-column='$colIndex' data-status='$dataStatus'>$text</td>`n"
        $colIndex++
    }
    $htmlBody += "                </tr>`n"
}

$htmlFooter = @"
            </tbody>
        </table>
    </div>

    <div class="footer">
        <p>Authentication Methods Report | MFA Health Score included</p>
    </div>
</div>

<script>
const columnMapping = {1:'Email',2:'External',3:'Fido2',4:'Password',5:'Authenticator',6:'AuthenticatorPasswordless',7:'HardwareOath',8:'Phone',9:'SoftwareOath',10:'TemporaryAccessPass',11:'WindowsHello',12:'PlatformCredential',13:'QRCodePIN'};

function parseDate(s){ if(!s || s==='Never' || s==='Registered (never used)') return null; const d=new Date(s); return isNaN(d)?null:d; }
function daysSince(s){ const d=parseDate(s); if(!d) return Infinity; const now=new Date(); return Math.floor((now-d)/(1000*60*60*24)); }
function parseCreated(s){ if(!s) return null; const d=new Date(s); return isNaN(d)?null:d; }

function getRows(){ return Array.from(document.querySelectorAll('#dataTable tbody tr')); }

function applyFilters(){
    const searchTerm  = document.getElementById('searchInput').value.toLowerCase();
    const inactivity  = document.getElementById('inactivityFilter').value;
    const methodName  = document.getElementById('methodFilter').value;
    const methodStat  = document.getElementById('methodStatusFilter').value;
    const protect     = document.getElementById('protectFido2Auth').checked;

    const sortField   = document.getElementById('sortField').value;
    const sortDir     = document.getElementById('sortDirection').value; // desc/asc
    const neverFirst  = document.getElementById('neverFirst').checked;

    const rows = getRows();
    let visible = 0;

    // Filter
    rows.forEach(row=>{
        let show = true;
        const cells = row.querySelectorAll('td');
        const upn = cells[0].textContent.toLowerCase();

        const created = parseCreated(cells[1].textContent);

        const hasAny       = (cells[2].textContent.trim().toLowerCase()==='true');
        const hasFido2     = (cells[4].textContent.trim().toLowerCase()==='true');
        const authNever    = (cells[10].textContent.trim().toLowerCase()==='true'); // col E in header, but index here is 10
        const baseDoNot    = (cells[11].textContent.trim().toLowerCase()==='true');
        const baseSafe     = (cells[12].textContent.trim().toLowerCase()==='true');

        const effectiveDoNot = protect ? baseDoNot : false;
        const effectiveSafe  = protect ? baseSafe  : (authNever && !effectiveDoNot);

        if (searchTerm && !upn.includes(searchTerm)) show=false;

        // inactivity
        if (show && inactivity){
            if (inactivity==='never'){
                let anyUsed=false;
                // methods start after meta: meta up to index 14 → methods start at 15
                for (let i=15;i<cells.length;i++){
                    if (cells[i].dataset.status === 'used') { anyUsed=true; break; }
                }
                if (anyUsed) show=false;
            } else {
                let mostRecent = Infinity;
                for (let i=15;i<cells.length;i++){
                    const days = daysSince(cells[i].textContent);
                    if (days < mostRecent) mostRecent = days;
                }
                if (mostRecent < parseInt(inactivity)) show=false;
            }
        }

        // method filter
        if (show && methodName && methodStat){
            let methodIdx = -1;
            for (const [idx,name] of Object.entries(columnMapping)){
                if (name===methodName){ methodIdx=parseInt(idx); break; }
            }
            if (methodIdx!==-1){
                const cell = cells[14 + methodIdx]; // shift: 0..14 meta, methods start at 15
                const status = cell.dataset.status; // used | never | registered-never
                if (methodStat==='has'   && status!=='used') show=false;
                if (methodStat==='never' && status==='used') show=false;
            }
        }

        if (show){ row.classList.remove('hidden-row'); visible++; }
        else { row.classList.add('hidden-row'); }
    });

    // Sort visible
    const visibleRows = rows.filter(r => !r.classList.contains('hidden-row'));
    if (sortField){
        const dir = (sortDir==='asc') ? 1 : -1;
        visibleRows.sort((a,b)=>{
            const ca = a.querySelectorAll('td');
            const cb = b.querySelectorAll('td');

            const createdA = parseCreated(ca[1].textContent);
            const createdB = parseCreated(cb[1].textContent);

            const mostRecentDays = (cells)=>{
                let m = Infinity;
                for (let i=15;i<cells.length;i++){
                    const d = daysSince(cells[i].textContent);
                    if (d < m) m = d;
                }
                return m;
            };

            const lastUsedFor = (cells, methodName)=>{
                let methodIdx = -1;
                for (const [idx,name] of Object.entries(columnMapping)){
                    if (name===methodName){ methodIdx=parseInt(idx); break; }
                }
                if (methodIdx===-1) return {status:'never',date:null};
                const c = cells[14 + methodIdx];
                const status = c.dataset.status;
                const d = parseDate(c.textContent);
                return {status, date:d};
            };

            let va, vb;

            if (sortField==='CreatedDate'){
                va = createdA ? createdA.getTime() : 0;
                vb = createdB ? createdB.getTime() : 0;
                return (va - vb) * dir;
            }
            if (sortField==='MostRecent'){
                va = mostRecentDays(ca);
                vb = mostRecentDays(cb);
                return (va - vb) * -1 * dir; // smaller days = more recent
            }
            if (sortField==='Authenticator' || sortField==='Fido2'){
                const ma = lastUsedFor(ca, sortField);
                const mb = lastUsedFor(cb, sortField);

                const aIsNever = (ma.status!=='used');
                const bIsNever = (mb.status!=='used');

                if (aIsNever || bIsNever){
                    if (aIsNever && bIsNever){
                        if (neverFirst){
                            const ta = createdA ? createdA.getTime() : 0;
                            const tb = createdB ? createdB.getTime() : 0;
                            return (ta - tb) * (dir);
                        } else { return 0; }
                    }
                    if (neverFirst){ return aIsNever ? -1 : 1; }
                    else { return aIsNever ? 1 : -1; }
                }

                const ta = ma.date ? ma.date.getTime() : 0;
                const tb = mb.date ? mb.date.getTime() : 0;
                return (ta - tb) * dir;
            }
            return 0;
        });

        const tbody = document.querySelector('#dataTable tbody');
        visibleRows.forEach(r => tbody.appendChild(r));
    }

    document.getElementById('visibleUsers').textContent = visible;
    toggleEmptyColumns();
}

function toggleEmptyColumns(){
    const table = document.getElementById('dataTable');
    const shouldHide = document.getElementById('hideEmptyColumns').checked;

    for (let col=1; col<=13; col++){
        let hasData=false;
        const rows = document.querySelectorAll('#dataTable tbody tr:not(.hidden-row)');
        rows.forEach(row=>{
            const cell = row.querySelector('td[data-column="'+col+'"]');
            if (!cell) return;
            if (cell.dataset && (cell.dataset.status==='used' || cell.dataset.status==='registered-never')) hasData=true;
        });

        const headers = table.querySelectorAll('th[data-column="'+col+'"]');
        const dataCells = table.querySelectorAll('td[data-column="'+col+'"]');

        if (shouldHide && !hasData){
            headers.forEach(h=>h.classList.add('hidden-column'));
            dataCells.forEach(c=>c.classList.add('hidden-column'));
        } else {
            headers.forEach(h=>h.classList.remove('hidden-column'));
            dataCells.forEach(c=>c.classList.remove('hidden-column'));
        }
    }
}

function exportVisibleToCsv(){
    const rows = Array.from(document.querySelectorAll('#dataTable tbody tr')).filter(r => !r.classList.contains('hidden-row'));
    if (rows.length===0){ alert('No visible rows to export.'); return; }

    // headers
    const headerCells = Array.from(document.querySelectorAll('#dataTable thead th'));
    const headers = headerCells.map(th => th.textContent.trim());

    const data = [headers];

    rows.forEach(row=>{
        const cells = Array.from(row.querySelectorAll('td'));
        const rowVals = cells.map(td => {
            let t = td.textContent.trim();
            // badge temizliği
            t = t.replace(/\s*FIDO2 present — keep Authenticator\s*/g,'')
                 .replace(/\s*Safe to remove Authenticator\s*/g,'');
            if (t.includes('"') || t.includes(',') || t.includes('\n')){
                t = '"' + t.replace(/"/g,'""') + '"';
            }
            return t;
        });
        data.push(rowVals);
    });

    const csv = data.map(r => r.join(',')).join('\n');
    const blob = new Blob([csv], {type:'text/csv;charset=utf-8;'});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    const now = new Date();
    const ts = now.toISOString().replace(/[:T]/g,'-').split('.')[0];
    a.href = url;
    a.download = 'AuthMethods_Visible_' + ts + '.csv';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

window.addEventListener('DOMContentLoaded', ()=>{ 
    // küçük animasyon: progress bar'ları doldur
    document.querySelectorAll('.progress > div').forEach(el=>{
        const w = el.style.width; el.style.width = '0%';
        setTimeout(()=>{ el.style.width = w; }, 50);
    });
    applyFilters(); 
});
</script>
</body>
</html>
"@

# Combine all HTML parts (OFFLINE — no web requests)
$htmlReport = $htmlHeader + $htmlBody + $htmlFooter

# ----------------------------
# Export flow (Prompt/HTML/CSV/Both)
# ----------------------------
Add-Type -AssemblyName System.Windows.Forms

function Save-File([string]$defaultName,[string]$filter){
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = $filter
    $dlg.Title = "Save Authentication Methods Report"
    $dlg.FileName = $defaultName
    $dlg.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    if ($dlg.ShowDialog() -eq 'OK'){ return $dlg.FileName } else { return $null }
}

if ($ExportFormat -eq "Prompt") {
    Write-Host "Export type? [H]TML, [C]SV, [B]oth" -ForegroundColor Cyan
    $choice = (Read-Host "Select H/C/B").ToUpper()
    switch ($choice) {
        "H" { $ExportFormat = "HTML" }
        "C" { $ExportFormat = "CSV" }
        "B" { $ExportFormat = "Both" }
        default { $ExportFormat="HTML" }
    }
}

if ($ExportFormat -in @("CSV","Both")) {
    $csvName = "AuthenticationMethodsReport_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss')
    $csvPath = Save-File -defaultName $csvName -filter "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    if ($csvPath) {
        $authMethodsReport | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $csvPath
        Write-Host "CSV exported: $csvPath" -ForegroundColor Green
    } else {
        Write-Host "CSV export skipped." -ForegroundColor Yellow
    }
}

if ($ExportFormat -in @("HTML","Both")) {
    $htmlName = "AuthenticationMethodsReport_{0}.html" -f (Get-Date -Format 'yyyyMMdd_HHmmss')
    $htmlPath = Save-File -defaultName $htmlName -filter "HTML Files (*.html)|*.html|All Files (*.*)|*.*"
    if ($htmlPath) {
        $htmlReport | Out-File -FilePath $htmlPath -Encoding UTF8
        Write-Host "HTML exported: $htmlPath" -ForegroundColor Green
        $open = Read-Host "Open HTML report now? (Y/N)"
        if ($open -match '^[Yy]$') { Start-Process $htmlPath }
    } else {
        Write-Host "HTML export skipped." -ForegroundColor Yellow
    }
}

Write-Host "Done." -ForegroundColor Cyan
