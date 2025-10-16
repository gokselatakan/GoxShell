$day = 90
$date = (Get-Date).AddDays(-$day)

# Find inactive users
$inactiveUsers = Get-ADUser -Filter {LastLogonDate -lt $date -and Enabled -eq $true} -Properties LastLogonDate, DistinguishedName, emailaddress |
    Where-Object {
        ($_.'DistinguishedName' -notmatch 'OU=Service Accounts') -and 
        ($_.'DistinguishedName' -notmatch 'OU=Shared Mailboxes') -and 
        ($_.'DistinguishedName' -notmatch 'OU=Resource Mailboxes') -and 
        ($_.emailaddress -notmatch 'HealthMailbox')
    } |
    Select-Object Name, SamAccountName, LastLogonDate, DistinguishedName, emailaddress |
    Sort-Object LastLogonDate -Descending

# Determine the number of days they have been inactive
function DaysInactive ([datetime]$LastLogonDate) {
    return ((Get-Date) - $LastLogonDate).Days
}

# Function to get OU information
function Get-OU([string]$DistinguishedName) {
    $ouParts = $DistinguishedName -split ',' | Where-Object { $_ -like "OU=*" }
    $ou = $ouParts -join ',' # Create the OU information
    return $ou
}

# Process user data and write to CSV file
$results = foreach ($user in $inactiveUsers) {
    $daysInactive = DaysInactive -LastLogonDate $user.LastLogonDate
    $ou = Get-OU -DistinguishedName $user.DistinguishedName

    [PSCustomObject]@{
        Name               = $user.Name
        SamAccountName     = $user.SamAccountName
        LastLogonDate      = $user.LastLogonDate
        DaysInactive       = $daysInactive
        OrganizationalUnit = $ou
        Mail               = $user.emailaddress
    }
}

# Export results to CSV file
$desktopPath = [Environment]::GetFolderPath('Desktop')
$results | Export-Csv -Path "$desktopPath\inactiveusers.csv" -NoTypeInformation -Encoding UTF8
