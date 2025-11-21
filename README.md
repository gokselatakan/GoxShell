PowerShell Automation Scripts for Azure AD, Entra ID & On-Prem AD

This repository contains a collection of PowerShell automation scripts designed to streamline and simplify administrative tasks across Azure AD / Entra ID, Microsoft 365, and On-Prem Active Directory environments.

These scripts support identity operations, group management, token revocation, MFA reporting, inactive user analysis, and general directory automation.

📂 Contents

Cloud account management automation

Group administration (Cloud & On-Prem)

User directory lookups & reporting

Inactive / dormant account investigations

MFA & authentication reporting

Session-based and activity-based reporting

🚀 Usage

Scripts are intended to be executed in environments with appropriate modules installed.
Most scripts rely on:

Microsoft Graph PowerShell

AzureAD module (legacy scenarios)

ActiveDirectory module (On-Prem AD)

MSOL module (optional legacy usage)

Before running any script:

Set-ExecutionPolicy RemoteSigned -Scope CurrentUser

Required modules:
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module AzureAD -Scope CurrentUser
Install-WindowsFeature RSAT-AD-PowerShell

🤝 Contributions

Contributions are welcome.
Feel free to submit improvements, enhanced error handling, new automation logic, or additional reporting scenarios via pull requests.
