# Step 1: Ensure the Exchange Online Management module is installed
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
}

# Step 2: Import the module
Import-Module ExchangeOnlineManagement

# Step 3: Connect to Exchange Online using modern authentication
# This will prompt you to sign in interactively with MFA support
Connect-ExchangeOnline

# Step 4: Run the OpenText ETP Rules Script directly from GitHub
Invoke-Expression (Invoke-WebRequest -Uri "https://raw.githubusercontent.com/josh-goipower/opentext-etp-rules-script/main/Create_OpenText_Email_Threat_Protection_Rules.ps1" -UseBasicParsing).Content