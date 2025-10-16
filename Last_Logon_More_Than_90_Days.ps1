# Import Active Directory Module
Import-Module ActiveDirectory

# Get the current date
$currentDate = Get-Date

# Define the time span of 3 months (approximately 90 days)
$idleDaysThreshold = 90

# Calculate the threshold date for last logon (more than 3 months ago)
$thresholdDate = $currentDate.AddDays(-$idleDaysThreshold)

# Initialize an array to hold the results
$allIdleUsers = @()

# Specify the distinguished names of the OUs to search within, if needed
# $searchBaseOUs = @("OU=Sites,DC=DOMAINNAME,DC=local", "OU=Admins,OU=Core,DC=DOMAINNAME,DC=local", "CN=Users,DC=DOMAINNAME,DC=local")

# Get all user accounts with last logon date more than 3 months ago
# Optionally, use -SearchBase parameter with each $searchBaseOU in a foreach loop if searching specific OUs
$allIdleUsers = Get-ADUser -Filter {LastLogonTimeStamp -lt $thresholdDate -and Enabled -eq $true} -Properties Name, SamAccountName, LastLogonTimeStamp -ResultSetSize $null

# Export the results to a CSV file
$exportPath = "C:\temp\IdleAccountsReport.csv"
$allIdleUsers | Select-Object Name, SamAccountName, @{Name='LastLogonDate'; Expression={[DateTime]::FromFileTime($_.LastLogonTimeStamp)}} | Export-Csv -Path $exportPath -NoTypeInformation

Write-Host "Exported idle-stale accounts report to $exportPath"
