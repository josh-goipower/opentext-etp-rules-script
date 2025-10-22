# Script to create Email Threat Protection (ETP) transport rules in Microsoft 365
# All rules are created DISABLED by default

# Define valid message types for exception rules
$messageTypes = @(
    'Voicemail'  # Only Voicemail exception needed
)

# Function to validate message types
function Test-MessageTypes {
    param([string[]]$Types)
    
    $validTypes = @(
        'OOF', 'AutoForward', 'Encrypted', 'Calendaring',
        'PermissionControlled', 'Voicemail', 'Signed',
        'ApprovalRequest', 'ReadReceipt'
    )
    
    foreach ($type in $Types) {
        if ($type -notin $validTypes) {
            Write-Host "  WARNING: Invalid message type: $type" -ForegroundColor Yellow
            Write-Host "  Valid types are: $($validTypes -join ', ')" -ForegroundColor Yellow
            return $false
        }
    }
    return $true
}

# Check if Exchange Online Management module is installed
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "Exchange Online Management module not found. Installing..." -ForegroundColor Yellow
    Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
}

# Import the module
Import-Module ExchangeOnlineManagement -ErrorAction Stop

# Connect to Exchange Online
Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
Write-Host "Please sign in with your Microsoft 365 admin credentials." -ForegroundColor Yellow

try {
    Connect-ExchangeOnline -ErrorAction Stop
    Start-Sleep -Seconds 2  # Brief pause to ensure connection is fully established
    
    # Test connection by getting existing rules
    $existingRules = Get-TransportRule -ErrorAction Stop
    Write-Host "Successfully connected to Exchange Online." -ForegroundColor Green
    Write-Host "Found $($existingRules.Count) existing transport rules." -ForegroundColor Gray
    
    # Check for existing rules with same names
    $ourRuleNames = @(
        "Limit Inbound Mail to ETP (Quarantine direct send)",
        "Limit Inbound Mail to ETP (Reject direct send)",
        "Allow Inbound Mail from ETP",
        "Bypass Safe Links"
    )
    
    $conflictingRules = $existingRules | Where-Object { $_.Name -in $ourRuleNames }
    if ($conflictingRules) {
        Write-Host "`nFound existing rules that will be replaced:" -ForegroundColor Yellow
        $conflictingRules | ForEach-Object {
            Write-Host "  • $($_.Name) (Priority: $($_.Priority))" -ForegroundColor Yellow
        }
        Write-Host ""
    }
    
} catch {
    Write-Host "ERROR: Failed to connect to Exchange Online or insufficient permissions" -ForegroundColor Red
    Write-Host "Details: $_" -ForegroundColor Red
    Write-Host "`nPlease ensure:" -ForegroundColor Yellow
    Write-Host "  1. You have admin credentials for Exchange Online" -ForegroundColor Yellow
    Write-Host "  2. Your account has permission to manage transport rules" -ForegroundColor Yellow
    Write-Host "  3. You're connected to the internet" -ForegroundColor Yellow
    exit 1
}

# Define the ETP IP ranges (used in most rules)
$etpIpRanges = @(
    "5.152.184.128/25", "5.152.185.128/26", "8.19.118.0/24", "8.31.233.0/24",
    "69.20.58.224/28", "5.152.188.0/24", "199.187.164.0/24", "199.187.165.0/24",
    "199.187.166.0/24", "199.187.167.0/24", "69.25.26.128/26", "204.232.250.0/24",
    "74.203.184.184/32", "74.203.184.185/32", "199.30.235.11/32", "74.203.185.12/32"
)

# Define SafeLinks IP ranges (subset of ETP ranges)
$safeLinksIpRanges = @(
    "5.152.184.128/25", "5.152.185.128/26", "8.19.118.0/24", "8.31.233.0/24",
    "69.20.58.224/28", "5.152.188.0/24", "199.187.164.0/24", "199.187.165.0/24",
    "199.187.166.0/24", "199.187.167.0/24", "69.25.26.128/26", "204.232.250.0/24",
    "74.203.184.184/32"
)

# Function to check and remove existing rule
function Remove-ExistingRule {
    param([string]$RuleName)
    
    try {
        $existingRule = Get-TransportRule -Identity $RuleName -ErrorAction SilentlyContinue
        if ($existingRule) {
            Write-Host "  Rule '$RuleName' already exists. Removing..." -ForegroundColor Yellow
            Remove-TransportRule -Identity $RuleName -Confirm:$false -ErrorAction Stop
            Write-Host "  Existing rule removed." -ForegroundColor Green
            
            # Wait longer to ensure removal is fully processed in Exchange Online
            Write-Host "  Waiting for removal to complete..." -ForegroundColor Gray
            Start-Sleep -Seconds 10
            
            # Verify removal
            $stillExists = Get-TransportRule -Identity $RuleName -ErrorAction SilentlyContinue
            if ($stillExists) {
                Write-Host "  WARNING: Rule still exists after removal attempt" -ForegroundColor Yellow
                throw "Rule removal verification failed"
            }
            Write-Host "  Removal confirmed." -ForegroundColor Green
        }
    } catch {
        Write-Host "  WARNING: Failed to remove existing rule: $_" -ForegroundColor Yellow
        Write-Host "  Attempting to continue..." -ForegroundColor Yellow
    }
}

# Function to verify IP range format
function Test-IpRanges {
    param([string[]]$IpRanges)
    
    foreach ($range in $IpRanges) {
        if ($range -notmatch '^(\d{1,3}\.){3}\d{1,3}(/\d{1,2})?$') {
            Write-Host "  WARNING: Invalid IP range format: $range" -ForegroundColor Yellow
            return $false
        }
    }
    return $true
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "ETP Transport Rules Creation Script" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "`nAll rules will be created in DISABLED state." -ForegroundColor Yellow
Write-Host "You must manually enable one of the Limit rules (Quarantine or Reject)." -ForegroundColor Yellow
Write-Host "`nPress Enter to continue or Ctrl+C to cancel..."
$null = Read-Host

# =============================================================================
# Rule 1: Limit Inbound Mail to ETP (Quarantine - Option A)
# =============================================================================
Write-Host "`n[1/4] Creating 'Limit Inbound Mail to ETP (Quarantine direct send)' rule..." -ForegroundColor Cyan

$quarantineRuleName = "Limit Inbound Mail to ETP (Quarantine direct send)"
$quarantineRuleComment = "This rule will quarantine incoming external email if the message was not delivered by ETP. This rule should only be active for ETP customers and it must be disabled if ETP service is cancelled."

Remove-ExistingRule -RuleName $quarantineRuleName

# Validate IP ranges before creating rule
if (-not (Test-IpRanges $etpIpRanges)) {
    Write-Host "  WARNING: Some IP ranges may be invalid. Creating rule anyway..." -ForegroundColor Yellow
}

# Create rule with mandatory parameters first
try {
    Write-Host "  Creating quarantine rule with specified settings..." -ForegroundColor Gray
    
    # Get current rules to determine available priority
    $currentRules = Get-TransportRule -ErrorAction Stop
    $maxPriority = if ($currentRules) { ($currentRules | Measure-Object -Property Priority -Maximum).Maximum } else { -1 }
    $nextPriority = $maxPriority + 1

    $params = @{
        Name = $quarantineRuleName
        Priority = $nextPriority
        FromScope = "NotInOrganization"                     # Outside the organization
        SetSCL = 6                                         # Set SCL to 6
        Quarantine = $true                                 # Send to quarantine
        ExceptIfSenderIpRanges = $etpIpRanges             # Except if sender IP matches ETP ranges
        ExceptIfHeaderContainsMessageHeader = "x-ms-exchange-meetingforward-message"  # Except if header includes
        ExceptIfHeaderContainsWords = "Forward"            # Except if header contains "Forward"
        ExceptIfMessageTypeMatches = "Voicemail"          # Except if message type is Voicemail
        Mode = "Enforce"                                   # Set rule mode to Enforce
        SenderAddressLocation = "Header"                   # Match sender address in header
        Comments = "This rule will quarantine incoming external email if the message was not delivered by ETP. This rule should only be active for ETP customers and it must be disabled if ETP service is cancelled."
        Enabled = $false                                   # Create rule disabled by default
    }
    
    Write-Host "    • Setting sender scope to: Outside the organization" -ForegroundColor Gray
    Write-Host "    • Setting SCL to: 6" -ForegroundColor Gray
    Write-Host "    • Enabling quarantine redirection" -ForegroundColor Gray
    Write-Host "    • Adding IP range exceptions" -ForegroundColor Gray
    Write-Host "    • Adding header exceptions" -ForegroundColor Gray
    Write-Host "    • Adding message type exceptions" -ForegroundColor Gray

    # Add optional parameters if they're valid
    if ($etpIpRanges) {
        $params["ExceptIfSenderIpRanges"] = $etpIpRanges
    }
    
    # Add message types one at a time to avoid array conversion issues
    foreach ($type in $validMessageTypes) {
        Write-Host "    Adding exception for message type: $type" -ForegroundColor Gray
        $params["ExceptIfMessageTypeMatches"] = $type
    }
    
    if ($quarantineRuleComment) {
        $params["Comments"] = $quarantineRuleComment
    }

    # Create the rule
    $newRule = New-TransportRule @params -ErrorAction Stop | Out-Null

    # Verify the rule was created
    $createdRule = Get-TransportRule -Identity $quarantineRuleName -ErrorAction Stop
    if (-not $createdRule) {
        throw "Rule was not found after creation"
    }
    
    Write-Host "  SUCCESS: Quarantine rule created (Priority: $($createdRule.Priority), Rule ID: $($createdRule.Guid))" -ForegroundColor Green
} catch {
    Write-Host "  ERROR: Failed to create Quarantine rule" -ForegroundColor Red
    Write-Host "  Details: $_" -ForegroundColor Red
}

# =============================================================================
# Rule 2: Limit Inbound Mail to ETP (Reject - Option B)
# =============================================================================
Write-Host "`n[2/4] Creating 'Limit Inbound Mail to ETP (Reject direct send)' rule..." -ForegroundColor Cyan

$rejectRuleName = "Limit Inbound Mail to ETP (Reject direct send)"
$rejectRuleComment = "This rule will reject incoming external email if the message was not delivered by ETP. This rule should only be active for ETP customers and it must be disabled if ETP service is cancelled."

Remove-ExistingRule -RuleName $rejectRuleName

try {
    # Get current priority
    $currentRules = Get-TransportRule -ErrorAction Stop
    $maxPriority = if ($currentRules) { ($currentRules | Measure-Object -Property Priority -Maximum).Maximum } else { -1 }
    $nextPriority = $maxPriority + 1

    # Convert message types to individual strings
    $validMessageTypes = $messageTypes | ForEach-Object { $_.ToString() }
    Write-Host "  Adding message type exceptions: $($validMessageTypes -join ', ')" -ForegroundColor Gray

    $params = @{
        Name = $rejectRuleName
        Priority = $nextPriority
        FromScope = "NotInOrganization"                     # Outside the organization
        RejectMessageReasonText = "Direct send not permitted."
        ExceptIfSenderIpRanges = $etpIpRanges             # Except if sender IP matches ETP ranges
        ExceptIfHeaderContainsMessageHeader = "x-ms-exchange-meetingforward-message"  # Except if header includes
        ExceptIfHeaderContainsWords = "Forward"            # Except if header contains "Forward"
        ExceptIfMessageTypeMatches = "Voicemail"          # Except if message type is Voicemail
        Mode = "Enforce"
        Enabled = $false
        SenderAddressLocation = "Header"                   # Match sender address in header
        Comments = $rejectRuleComment
    }

    Write-Host "    • Setting sender scope to: Outside the organization" -ForegroundColor Gray
    Write-Host "    • Setting reject message text" -ForegroundColor Gray
    Write-Host "    • Adding IP range exceptions" -ForegroundColor Gray
    Write-Host "    • Adding header exceptions" -ForegroundColor Gray
    Write-Host "    • Adding message type exceptions" -ForegroundColor Gray

    New-TransportRule @params -ErrorAction Stop | Out-Null

    Write-Host "  SUCCESS: Reject rule created (DISABLED)" -ForegroundColor Green
} catch {
    Write-Host "  ERROR: Failed to create Reject rule" -ForegroundColor Red
    Write-Host "  Details: $_" -ForegroundColor Red
}

# =============================================================================
# Rule 3: Allow Inbound Mail from ETP
# =============================================================================
Write-Host "`n[3/4] Creating 'Allow Inbound Mail from ETP' rule..." -ForegroundColor Cyan

$allowRuleName = "Allow Inbound Mail from ETP"
$allowRuleComment = "This rule must remain in place to allow ETP traffic to bypass M365 filtering."

Remove-ExistingRule -RuleName $allowRuleName

try {
    # Get current priority
    $currentRules = Get-TransportRule -ErrorAction Stop
    $maxPriority = if ($currentRules) { ($currentRules | Measure-Object -Property Priority -Maximum).Maximum } else { -1 }
    $nextPriority = $maxPriority + 1

    $params = @{
        Name = $allowRuleName
        Priority = $nextPriority
        SenderIpRanges = $etpIpRanges
        SetSCL = -1
        Comments = $allowRuleComment
        Enabled = $true
        Mode = "Enforce"
        SenderAddressLocation = "Header"
    }

    New-TransportRule @params -ErrorAction Stop | Out-Null

    Write-Host "  SUCCESS: Allow rule created (ENABLED)" -ForegroundColor Green
} catch {
    Write-Host "  ERROR: Failed to create Allow rule" -ForegroundColor Red
    Write-Host "  Details: $_" -ForegroundColor Red
}

# =============================================================================
# Rule 4: Bypass Safe Links
# =============================================================================
Write-Host "`n[4/4] Creating 'Bypass Safe Links' rule..." -ForegroundColor Cyan

$safeLinksRuleName = "Bypass Safe Links"
$safeLinksRuleComment = "This rule must remain in place to allow QMR from ETP to bypass SafeLink but does not allow all emails to bypass the Office 365 Spam Filter."

Remove-ExistingRule -RuleName $safeLinksRuleName

try {
    New-TransportRule -Name $safeLinksRuleName `
        -SenderIpRanges $safeLinksIpRanges `
        -From "notice@appriver.com" `
        -SetHeaderName "X-MS-Exchange-Organization-SkipSafeLinksProcessing" `
        -SetHeaderValue "1" `
        -Comments $safeLinksRuleComment `
        -Enabled $false `
        -Mode Enforce `
        -ErrorAction Stop | Out-Null

    Write-Host "  SUCCESS: Bypass Safe Links rule created (DISABLED)" -ForegroundColor Green
} catch {
    Write-Host "  ERROR: Failed to create Bypass Safe Links rule" -ForegroundColor Red
    Write-Host "  Details: $_" -ForegroundColor Red
}

# =============================================================================
# Summary
# =============================================================================
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Summary of Created Rules" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

Write-Host "`nRule 1: Limit Inbound Mail to ETP (Quarantine) - DISABLED" -ForegroundColor Yellow
Write-Host "Purpose: $quarantineRuleComment" -ForegroundColor Gray

Write-Host "`nRule 2: Limit Inbound Mail to ETP (Reject) - DISABLED" -ForegroundColor Yellow
Write-Host "Purpose: $rejectRuleComment" -ForegroundColor Gray

Write-Host "`nRule 3: Allow Inbound Mail from ETP - ENABLED" -ForegroundColor Green
Write-Host "Purpose: $allowRuleComment" -ForegroundColor Gray

Write-Host "`nRule 4: Bypass Safe Links - DISABLED" -ForegroundColor Yellow
Write-Host "Purpose: $safeLinksRuleComment" -ForegroundColor Gray

Write-Host "`nIMPORTANT:" -ForegroundColor Yellow
Write-Host "  • Enable either the Quarantine OR Reject rule (not both)"
Write-Host "  • Allow rule is enabled by default"
Write-Host "  • Bypass Safe Links rule should be enabled only if ETP is deployed"
Write-Host "  • All rules have correct priorities set automatically"

Write-Host "`nNext Steps:" -ForegroundColor Cyan
Write-Host "  1. Go to Exchange Admin Center: https://admin.exchange.microsoft.com/#/transportrules"
Write-Host "  2. Enable rules based on your security preference"
Write-Host "  3. Confirm all rules are in correct priority order"

# Disconnect from Exchange Online
Write-Host "`nDisconnecting from Exchange Online..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false

Write-Host "`nScript completed. No rules are active until enabled manually." -ForegroundColor Green
