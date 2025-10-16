Add-Type -AssemblyName System.Windows.Forms

# ===================== Blue GUI for Inputs =====================
$form = New-Object System.Windows.Forms.Form
$form.Text = "SessionId-based Activity Export (M365)"
$form.Size = New-Object System.Drawing.Size(520, 300)
$form.BackColor = 'Blue'
$form.StartPosition = 'CenterScreen'

$lblUser = New-Object System.Windows.Forms.Label
$lblUser.Text = "User UPN (e.g., john@contoso.com):"
$lblUser.ForeColor = 'White'
$lblUser.AutoSize = $true
$lblUser.Location = New-Object System.Drawing.Point(10, 15)
$form.Controls.Add($lblUser)

$txtUser = New-Object System.Windows.Forms.TextBox
$txtUser.Size = New-Object System.Drawing.Size(470, 20)
$txtUser.Location = New-Object System.Drawing.Point(10, 35)
$form.Controls.Add($txtUser)

$lblSession = New-Object System.Windows.Forms.Label
$lblSession.Text = "Session Id (e.g., 007a6809-xxxx-xxxx-xxxx-5bafb8ab1740):"
$lblSession.ForeColor = 'White'
$lblSession.AutoSize = $true
$lblSession.Location = New-Object System.Drawing.Point(10, 65)
$form.Controls.Add($lblSession)

$txtSession = New-Object System.Windows.Forms.TextBox
$txtSession.Size = New-Object System.Drawing.Size(470, 20)
$txtSession.Location = New-Object System.Drawing.Point(10, 85)
$form.Controls.Add($txtSession)

$lblStart = New-Object System.Windows.Forms.Label
$lblStart.Text = "Start Date (optional, e.g., 12/15/2023):"
$lblStart.ForeColor = 'White'
$lblStart.AutoSize = $true
$lblStart.Location = New-Object System.Drawing.Point(10, 115)
$form.Controls.Add($lblStart)

$txtStart = New-Object System.Windows.Forms.TextBox
$txtStart.Size = New-Object System.Drawing.Size(220, 20)
$txtStart.Location = New-Object System.Drawing.Point(10, 135)
$form.Controls.Add($txtStart)

$lblEnd = New-Object System.Windows.Forms.Label
$lblEnd.Text = "End Date (optional, e.g., 12/31/2023):"
$lblEnd.ForeColor = 'White'
$lblEnd.AutoSize = $true
$lblEnd.Location = New-Object System.Drawing.Point(260, 115)
$form.Controls.Add($lblEnd)

$txtEnd = New-Object System.Windows.Forms.TextBox
$txtEnd.Size = New-Object System.Drawing.Size(220, 20)
$txtEnd.Location = New-Object System.Drawing.Point(260, 135)
$form.Controls.Add($txtEnd)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Run Export"
$btnRun.Location = New-Object System.Drawing.Point(200, 180)
$btnRun.Add_Click({ $form.Close() })
$form.Controls.Add($btnRun)

$form.ShowDialog()

# Read GUI inputs
$UserId    = $txtUser.Text.Trim()
$SessionId = $txtSession.Text.Trim()
$StartDateText = $txtStart.Text.Trim()
$EndDateText   = $txtEnd.Text.Trim()

# ===================== Prereqs: EXO module =====================
$Module = Get-Module ExchangeOnlineManagement -ListAvailable
if ($Module.Count -eq 0) {
  Write-Host "Exchange Online PowerShell module is not available." -ForegroundColor Yellow
  $Confirm = Read-Host "Install module now? [Y] Yes [N] No"
  if ($Confirm -match '^[yY]$') {
    Write-Host "Installing Exchange Online PowerShell module..."
    Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
  } else {
    Write-Host "EXO module is required. Please install: Install-Module ExchangeOnlineManagement" -ForegroundColor Red
    return
  }
}

# ===================== Connect to EXO =====================
Write-Host "`n🔄 Connecting to Exchange Online..." -ForegroundColor Cyan
Connect-ExchangeOnline -ShowBanner:$false

# ===================== Date Range Handling =====================
$MaxStartDate = ((Get-Date).AddDays(-179)).Date
$StartDate = $null
$EndDate   = $null

# Parse optional Start/End from GUI; else default to last 180 days
if ([string]::IsNullOrWhiteSpace($StartDateText)) {
  $StartDate = $MaxStartDate
} else {
  try {
    $StartDate = [DateTime]$StartDateText
    if ($StartDate -lt $MaxStartDate) {
      Write-Host "`nAudit can be retrieved only for the past 180 days. Choose a date after $MaxStartDate" -ForegroundColor Red
      Disconnect-ExchangeOnline -Confirm:$false | Out-Null
      return
    }
  } catch {
    Write-Host "`nNot a valid start date." -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    return
  }
}

if ([string]::IsNullOrWhiteSpace($EndDateText)) {
  $EndDate = (Get-Date).Date
} else {
  try {
    $EndDate = [DateTime]$EndDateText
    if ($EndDate -lt $StartDate) {
      Write-Host "End time should be later than start time." -ForegroundColor Red
      Disconnect-ExchangeOnline -Confirm:$false | Out-Null
      return
    }
  } catch {
    Write-Host "`nNot a valid end date." -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    return
  }
}

# Validate required inputs
if ([string]::IsNullOrWhiteSpace($UserId)) {
  Write-Host "User UPN is required." -ForegroundColor Red
  Disconnect-ExchangeOnline -Confirm:$false | Out-Null
  return
}
if ([string]::IsNullOrWhiteSpace($SessionId)) {
  Write-Host "Session Id is required." -ForegroundColor Red
  Disconnect-ExchangeOnline -Confirm:$false | Out-Null
  return
}

# ===================== Prep Output & Progress =====================
$Location = Get-Location
$OutputCSV = "$Location\SessionId_based-Activity_Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"

$IntervalTimeInMinutes = 1440
$CurrentStart = $StartDate
$CurrentEnd   = $CurrentStart.AddMinutes($IntervalTimeInMinutes)
if ($CurrentEnd -gt $EndDate) { $CurrentEnd = $EndDate }

$AggregateResultCount = 0
$i = 0

Write-Host "`n🟡 In progress: Retrieving sessionId-based activity log from $StartDate to $EndDate..." -ForegroundColor Yellow

# ===================== Main Loop =====================
while ($true) {
  if ($CurrentStart -eq $CurrentEnd) {
    Write-Host "Start and end time are the same. Please choose a wider range." -ForegroundColor Red
    break
  }

  $Results = Search-UnifiedAuditLog -StartDate $CurrentStart -EndDate $CurrentEnd `
             -UserIds $UserId -FreeText $SessionId -SessionId s `
             -SessionCommand ReturnLargeSet -ResultSize 5000

  $ResultCount = ($Results | Measure-Object).Count

  foreach ($Result in $Results) {
    $i++
    $AuditData  = $Result.AuditData | ConvertFrom-Json
    $ActivityTime = Get-Date($AuditData.CreationTime) -format g
    $UserName     = $AuditData.UserId
    $Operation    = $AuditData.Operation
    $ResultStatus = $AuditData.ResultStatus
    $Workload     = $AuditData.Workload

    [PSCustomObject]@{
      'Activity Time' = $ActivityTime
      'User Name'     = $UserName
      'Operation'     = $Operation
      'Result'        = $ResultStatus
      'Workload'      = $Workload
      'More Info'     = $Result.AuditData
    } | Select-Object 'Activity Time','User Name','Operation','Result','Workload','More Info' `
      | Export-Csv -Path $OutputCSV -NoTypeInformation -Append
  }

  Write-Progress -Activity "Retrieving audit log from $StartDate to $EndDate..." `
                 -Status "Processed records: $i" `
                 -PercentComplete ((($CurrentEnd - $StartDate).TotalMinutes / ($EndDate - $StartDate).TotalMinutes) * 100)

  $AggregateResultCount += $ResultCount

  if ($Results.Count -lt 5000) {
    if ($CurrentEnd -eq $EndDate) { break }
    $CurrentStart = $CurrentEnd
    if ($CurrentStart -gt (Get-Date)) { break }
    $CurrentEnd = $CurrentStart.AddMinutes($IntervalTimeInMinutes)
    if ($CurrentEnd -gt $EndDate) { $CurrentEnd = $EndDate }
  }
}

# ===================== Wrap Up =====================
if ($AggregateResultCount -eq 0) {
  Write-Host "`nNo audit records found for the given inputs." -ForegroundColor Yellow
} else {
  if (Test-Path -Path $OutputCSV) {
    Write-Host "`n✅ Exported $AggregateResultCount records." -ForegroundColor Green
    Write-Host "📄 File saved at: $OutputCSV" -ForegroundColor Cyan
  } else {
    Write-Host "`n❌ Export failed." -ForegroundColor Red
  }
}

Disconnect-ExchangeOnline -Confirm:$false | Out-Null
