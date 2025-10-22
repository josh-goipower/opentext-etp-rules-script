# OpenText Email Threat Protection (ETP) Rules Script# Robocopy Migration Tool



This PowerShell script automates the creation and configuration of Microsoft 365 transport rules required for OpenText Email Threat Protection (ETP) integration.A PowerShell-based wrapper for Robocopy that provides structured migration workflows with progress tracking, email notifications, and detailed reporting.



## Overview## Features



The script creates four essential transport rules in Microsoft 365:- **Structured Migration Workflow**

  - SEED: Initial copy with existing files skipped

1. **Limit Inbound Mail to ETP (Quarantine)** - Quarantines external emails not processed by ETP  - SYNC: Update changed files while preserving existing

   - Exceptions for Voicemail, meeting forwards, and ETP IP ranges  - RECONCILE: Back-sync newer files from destination to source

   - Created disabled by default  - MIRROR: Full synchronization (with safeguards)



2. **Limit Inbound Mail to ETP (Reject)** - Rejects external emails not processed by ETP- **Advanced Features**

   - Same exceptions as Quarantine rule  - Real-time progress tracking

   - Created disabled by default  - Detailed logging

  - Email notifications

3. **Allow Inbound Mail from ETP** - Allows mail from ETP IP ranges to bypass filtering  - Backup mode support

   - Created enabled by default  - Multi-threaded operations

   - Essential for ETP operation  - VSS fallback for locked files

  - Bandwidth throttling

4. **Bypass Safe Links** - Configures SafeLinks bypass for ETP QMR messages  - ACL preservation options

   - Created disabled by default

   - Enable only when ETP is deployed- **Safety Features**

  - Dry-run preview mode

## Prerequisites  - Mirror mode safeguards

  - Write access verification

- Windows PowerShell 5.1 or later  - Administrator privilege detection

- Exchange Online Management module  - Watchdog for stalled operations

- Microsoft 365 Global Admin or Exchange Admin privileges

- Network connectivity to Exchange Online## Requirements



## Installation- Windows OS (Windows 7/Server 2008 R2 or later)

- PowerShell 5.1 or later

1. Download the script:- Administrator rights (recommended)

   ```powershell

   git clone https://github.com/josh-goipower/opentext-etp-rules-script.git## Usage

   ```

### Basic Commands

2. Navigate to the script directory:

   ```powershell```powershell

   cd opentext-etp-rules-script# Initial copy (dry run)

   ```.\Robocopy_v6.ps1 -Mode SEED -Preview



## Usage# Perform initial copy

.\Robocopy_v6.ps1 -Mode SEED

1. Run the script:

   ```powershell# Sync changes

   .\Create_OpenText_Email_Threat_Protection_Rules.ps1.\Robocopy_v6.ps1 -Mode SYNC

   ```

# Back-sync newer files from destination

2. Follow the authentication prompts for Exchange Online.\Robocopy_v6.ps1 -Mode RECONCILE



3. The script will:# Mirror source to destination (requires confirmation)

   - Create all required rules with proper priorities.\Robocopy_v6.ps1 -Mode MIRROR -Confirm

   - Enable the Allow rule```

   - Leave other rules disabled for manual activation

### Common Options

4. Post-installation:

   - Enable either Quarantine OR Reject rule (not both)- `-Preview`: Dry-run mode (no changes made)

   - Enable Bypass Safe Links rule if using ETP- `-PrintCommand`: Display the constructed Robocopy command

   - Verify rule priorities in Exchange Admin Center- `-ForceBackup`: Use backup mode (/B) instead of restart mode (/Z)

- `-AppendLogs`: Append to existing log files instead of creating new ones

## Rule Details- `-VssFallback`: Try VSS snapshot if files are locked

- `-Confirm`: Required for MIRROR mode execution

### Limit Inbound Mail to ETP (Quarantine)

- Purpose: Quarantine external mail not processed by ETP### Configuration

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