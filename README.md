# OpenText ETP Transport Rules# OpenText Email Threat Protection Transport Rules Script# Microsoft 365 Transport Rules for OpenText ETP# OpenText Email Threat Protection (ETP) Rules Script# Robocopy Migration Tool



Automate Microsoft 365 transport rule setup for OpenText Email Threat Protection.



## What It DoesThis PowerShell script automates the creation of Microsoft 365 transport rules required for OpenText Email Threat Protection (ETP) service. It creates a set of rules to manage email flow through ETP while maintaining security and ensuring proper message handling.



This script creates four transport rules in Microsoft 365 to integrate with OpenText ETP:



**Quarantine Rule** - Quarantines external emails that bypass ETP  ## Transport Rules CreatedThis PowerShell script automates the creation of Microsoft 365 transport rules required for OpenText Email Threat Protection (ETP) service integration.

**Reject Rule** - Alternative to quarantine, rejects bypassed emails  

**Allow Rule** - Lets ETP-processed mail through (auto-enabled)  

**SafeLinks Bypass** - Allows ETP quarantine reports to skip SafeLinks

The script creates four essential rules:

## Quick Start



```powershell

.\Create_OpenText_Email_Threat_Protection_Rules.ps1### 1. Limit Inbound Mail to ETP (Quarantine direct send)## FeaturesThis PowerShell script automates the creation and configuration of Microsoft 365 transport rules required for OpenText Email Threat Protection (ETP) integration.A PowerShell-based wrapper for Robocopy that provides structured migration workflows with progress tracking, email notifications, and detailed reporting.

```

- **Status**: Created Disabled

The script will:

- Install required modules if needed- **Purpose**: Quarantines external emails not processed by ETP

- Connect to Exchange Online

- Create all four rules with proper settings- **Exceptions**:

- Leave enforcement rules disabled for your review

  - Voicemail messages### Created Rules

## After Running

  - Meeting forwards

1. Go to [Exchange Admin Center](https://admin.exchange.microsoft.com/#/transportrules)

2. Enable **either** Quarantine or Reject (not both)  - ETP IP ranges

3. Enable SafeLinks Bypass if using ETP

4. Verify rule priorities  - Forward headers



## What's Protected- **Settings**:1. **Limit Inbound Mail to ETP (Quarantine)**## Overview## Features



All rules automatically exclude:  - SCL set to 6

- Voicemail messages

- Meeting forwards  - Applies to external senders   - Quarantines external emails not processed by ETP

- ETP IP ranges

  - Quarantine action enabled

## Requirements

   - Created disabled by default

- PowerShell 5.1+

- Exchange admin permissions### 2. Limit Inbound Mail to ETP (Reject direct send)

- Internet connection

- **Status**: Created Disabled   - Includes exceptions for:

## License

- **Purpose**: Alternative to quarantine rule, rejects unprocessed mail

MIT
- **Exceptions**: Same as Quarantine rule     - Voicemail messagesThe script creates four essential transport rules in Microsoft 365:- **Structured Migration Workflow**

- **Settings**:

  - Rejection message configured     - Meeting forwards

  - Applies to external senders

     - ETP IP ranges  - SEED: Initial copy with existing files skipped

### 3. Allow Inbound Mail from ETP

- **Status**: Created Enabled

- **Purpose**: Allows ETP-processed mail to bypass filtering

- **Settings**:2. **Limit Inbound Mail to ETP (Reject)**1. **Limit Inbound Mail to ETP (Quarantine)** - Quarantines external emails not processed by ETP  - SYNC: Update changed files while preserving existing

  - SCL set to -1

  - Matches ETP IP ranges   - Alternative to Quarantine rule

  - Always enabled by default

   - Rejects external emails not processed by ETP   - Exceptions for Voicemail, meeting forwards, and ETP IP ranges  - RECONCILE: Back-sync newer files from destination to source

### 4. Bypass Safe Links

- **Status**: Created Disabled   - Created disabled by default

- **Purpose**: Allows QMR messages to bypass SafeLinks

- **Settings**:   - Includes same exceptions as Quarantine rule   - Created disabled by default  - MIRROR: Full synchronization (with safeguards)

  - Specific to notice@appriver.com

  - Only for ETP IP ranges

  - Sets SkipSafeLinksProcessing header

3. **Allow Inbound Mail from ETP**

## Prerequisites

   - Permits mail from ETP IP ranges to bypass M365 filtering

- Windows PowerShell 5.1 or higher

- Exchange Online Management module   - Created enabled by default2. **Limit Inbound Mail to ETP (Reject)** - Rejects external emails not processed by ETP- **Advanced Features**

- Microsoft 365 administrator credentials with transport rule management permissions

- Network connectivity to Exchange Online   - Essential for ETP operation



## Usage   - Same exceptions as Quarantine rule  - Real-time progress tracking



1. Run the script with administrator privileges:4. **Bypass Safe Links**

   ```powershell

   .\Create_OpenText_Email_Threat_Protection_Rules.ps1   - Allows QMR messages from ETP to bypass SafeLinks   - Created disabled by default  - Detailed logging

   ```

   - Created disabled by default

2. Follow the authentication prompts for Exchange Online

   - Specific to notice@appriver.com sender  - Email notifications

3. The script will:

   - Check for and install required modules

   - Connect to Exchange Online

   - Remove any existing rules with matching names### Key Capabilities3. **Allow Inbound Mail from ETP** - Allows mail from ETP IP ranges to bypass filtering  - Backup mode support

   - Create new rules with proper priorities

   - Enable the Allow rule

   - Leave other rules disabled for manual activation

- Automatic priority management   - Created enabled by default  - Multi-threaded operations

## Post-Installation Steps

- Comprehensive exception handling

1. Access the Exchange Admin Center:

   ```- IP range validation   - Essential for ETP operation  - VSS fallback for locked files

   https://admin.exchange.microsoft.com/#/transportrules

   ```- Rule removal verification



2. Choose your enforcement method:- Detailed progress logging  - Bandwidth throttling

   - Enable either Quarantine OR Reject rule (not both)

   - Verify Allow rule is enabled

   - Enable Bypass Safe Links if using ETP

## Requirements4. **Bypass Safe Links** - Configures SafeLinks bypass for ETP QMR messages  - ACL preservation options

3. Verify rule priorities are correct



## IP Range Configuration

- Windows PowerShell 5.1+   - Created disabled by default

The script includes predefined IP ranges for:

- ETP service- Exchange Online Management module

- SafeLinks processing

- Microsoft 365 admin credentials   - Enable only when ETP is deployed- **Safety Features**

These ranges are used to identify legitimate ETP traffic and allow proper message routing.

- Exchange administrator privileges

## Security Features

  - Dry-run preview mode

- Automatic priority management

- Rule creation verification## Quick Start

- IP range validation

- Comprehensive error handling## Prerequisites  - Mirror mode safeguards

- Safe default states (disabled)

1. **Download**

## Troubleshooting

   ```powershell  - Write access verification

If rule creation fails:

1. Verify Exchange Online connectivity   git clone https://github.com/josh-goipower/opentext-etp-rules-script.git

2. Check administrator permissions

3. Review any error messages in the console   cd opentext-etp-rules-script- Windows PowerShell 5.1 or later  - Administrator privilege detection

4. Ensure no conflicting rules exist

   ```

## Support

- Exchange Online Management module  - Watchdog for stalled operations

For issues or questions:

1. Verify script prerequisites are met2. **Run**

2. Check Exchange Online connection status

3. Review console output for any errors   ```powershell- Microsoft 365 Global Admin or Exchange Admin privileges

4. Ensure proper permissions are in place

   .\Create_OpenText_Email_Threat_Protection_Rules.ps1

## License

   ```- Network connectivity to Exchange Online## Requirements

MIT License - See LICENSE file for details


3. **Post-Setup**

   - Enable either Quarantine OR Reject rule (not both)

   - Allow rule is enabled automatically## Installation- Windows OS (Windows 7/Server 2008 R2 or later)

   - Enable Bypass Safe Links if using ETP

   - Verify rules in Exchange Admin Center- PowerShell 5.1 or later



## Configuration Details1. Download the script:- Administrator rights (recommended)



### Quarantine/Reject Rules   ```powershell

- Scope: External senders only

- Action: Quarantine/Reject non-ETP mail   git clone https://github.com/josh-goipower/opentext-etp-rules-script.git## Usage

- Exceptions:

  - ETP IP ranges   ```

  - Voicemail messages

  - Meeting forwards### Basic Commands

  - Forward headers

2. Navigate to the script directory:

### Allow Rule

- Purpose: Bypass M365 filtering for ETP   ```powershell```powershell

- Action: Sets SCL to -1

- Conditions: Matches ETP IP ranges   cd opentext-etp-rules-script# Initial copy (dry run)

- Status: Enabled by default

   ```.\Robocopy_v6.ps1 -Mode SEED -Preview

### Safe Links Rule

- Purpose: QMR message handling

- Sender: notice@appriver.com

- Action: Bypasses SafeLinks processing## Usage# Perform initial copy

- Status: Disabled by default

.\Robocopy_v6.ps1 -Mode SEED

## Security Considerations

1. Run the script:

- Rules created disabled for safety

- Automatic rule priority management   ```powershell# Sync changes

- IP range validation

- Comprehensive error handling   .\Create_OpenText_Email_Threat_Protection_Rules.ps1.\Robocopy_v6.ps1 -Mode SYNC

- Rule creation verification

   ```

## Troubleshooting

# Back-sync newer files from destination

1. **Rule Creation Issues**

   - Verify Exchange Online connection2. Follow the authentication prompts for Exchange Online.\Robocopy_v6.ps1 -Mode RECONCILE

   - Check admin privileges

   - Review error messages in console



2. **Post-Creation Steps**3. The script will:# Mirror source to destination (requires confirmation)

   - Access Exchange Admin Center

   - Verify rule priorities   - Create all required rules with proper priorities.\Robocopy_v6.ps1 -Mode MIRROR -Confirm

   - Enable appropriate rules

   - Test mail flow   - Enable the Allow rule```



3. **Common Issues**   - Leave other rules disabled for manual activation

   - Connection timeout: Check network

   - Permission errors: Verify admin rights### Common Options

   - Rule conflicts: Check existing rules

4. Post-installation:

## Support

   - Enable either Quarantine OR Reject rule (not both)- `-Preview`: Dry-run mode (no changes made)

For technical support:

1. Verify Exchange Online connectivity   - Enable Bypass Safe Links rule if using ETP- `-PrintCommand`: Display the constructed Robocopy command

2. Check script execution policy

3. Review console output for errors   - Verify rule priorities in Exchange Admin Center- `-ForceBackup`: Use backup mode (/B) instead of restart mode (/Z)

4. Submit issues via GitHub

- `-AppendLogs`: Append to existing log files instead of creating new ones

## Contributing

## Rule Details- `-VssFallback`: Try VSS snapshot if files are locked

1. Fork the repository

2. Create a feature branch- `-Confirm`: Required for MIRROR mode execution

3. Submit a pull request

### Limit Inbound Mail to ETP (Quarantine)

## License

- Purpose: Quarantine external mail not processed by ETP### Configuration

This project is licensed under the MIT License - see the LICENSE file for details.
- Exceptions:

  - ETP IP rangesEdit the `$Config` hashtable in the script to set:

  - Voicemail messages

  - Meeting forwards```powershell

  - Forward headers$Config = @{

    Source              = 'C:\RC-Source'              # Source path

### Limit Inbound Mail to ETP (Reject)    Destination         = 'C:\RC-Destination'         # Destination path

- Purpose: Reject external mail not processed by ETP    LogRoot            = 'C:\Logs\RC_Test'           # Log directory

- Exceptions: Same as Quarantine rule    Threads            = 32                          # Number of threads

- Alternative to Quarantine rule    PreserveACL        = $true                      # Copy security settings

    BandwidthThrottle  = 0                          # Milliseconds between packets

### Allow Inbound Mail from ETP    

- Purpose: Bypass M365 filtering for ETP traffic    # Exclude patterns

- Enabled by default    ExcludeDirs        = @()                        # e.g., @('Temp', '~snapshot')

- Required for ETP operation    ExcludeFiles       = @()                        # e.g., @('*.tmp', '*.bak')

    

### Bypass Safe Links    # Email notifications

- Purpose: Allow QMR from ETP to bypass SafeLinks    SendEmail          = $false

- Specific to notice@appriver.com    SmtpServer         = 'smtp.company.com'

- Enable after ETP deployment    EmailFrom          = 'migrations@company.com'

    EmailTo            = @('admin@company.com')

## Security Notes    SmtpPort          = 25

    SmtpUseSsl        = $false

- Script creates rules in disabled state for safety    

- Implements proper exception handling    # Watchdog settings

- Verifies rule creation and removal    WatchdogEnabled    = $true

- Validates IP ranges and message types    WatchdogTimeoutMin = 15                         # Minutes before timeout

}

## Support```



For issues or questions:## Migration Workflow

1. Check Exchange Admin Center for rule status

2. Verify Exchange Online connection1. **SEED Mode**

3. Check rule priorities and exceptions   - Initial copy operation

4. Submit an issue on GitHub   - Skips existing files at destination

   - Use `-Preview` first to see what will be copied

## Contributing

2. **SYNC Mode**

1. Fork the repository   - Updates changed files

2. Create a feature branch   - Preserves existing files at destination

3. Submit a pull request   - Good for incremental updates



## License3. **RECONCILE Mode**

   - Copies newer files from destination back to source

This project is licensed under the MIT License - see the LICENSE file for details   - Prevents data loss before mirror
   - Required before MIRROR mode

4. **MIRROR Mode**
   - Full synchronization
   - Deletes files at destination not in source
   - Requires `-Confirm` flag
   - Recommended to run RECONCILE first

## Logging and Monitoring

- Logs are stored in the configured `LogRoot` directory
- Each operation creates a timestamped subfolder
- Detailed statistics are displayed after each run
- Email notifications can be configured for success/failure

## Advanced Usage

### Backup Mode

```powershell
# Force backup mode
.\Robocopy_v6.ps1 -Mode SYNC -ForceBackup

# Use VSS fallback for locked files
.\Robocopy_v6.ps1 -Mode SYNC -VssFallback
```

### Log Management

```powershell
# Append to existing logs
.\Robocopy_v6.ps1 -Mode SYNC -AppendLogs
```

### Command Preview

```powershell
# See the exact Robocopy command that will be used
.\Robocopy_v6.ps1 -Mode SEED -Preview -PrintCommand
```

## Error Handling

- Exit codes follow Robocopy conventions:
  - 0: No changes needed
  - 1: Files copied successfully
  - 2: Extra files/directories detected
  - 3: Files copied, extras detected
  - 4-7: Mismatches present
  - ≥8: At least one failure occurred

## Best Practices

1. Always run with `-Preview` first
2. Use RECONCILE before MIRROR
3. Monitor logs for unexpected behaviors
4. Configure email notifications for unattended operations
5. Run with administrator rights when copying system files
6. Set appropriate thread count for your environment

## Troubleshooting

- Check log files for detailed error information
- Verify write permissions on destination
- Ensure network connectivity for remote paths
- Monitor system resources during large transfers
- Use `-PrintCommand` to verify Robocopy parameters

## Safety Features

The script includes several safety measures:

- Prevents MIRROR without confirmation
- Validates paths before copying
- Tests write access before operations
- Detects stalled operations
- Verifies administrator status