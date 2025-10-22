# OpenText ETP Transport Rules

Automate Microsoft 365 transport rule setup for OpenText Email Threat Protection.

## What It Does

Creates four transport rules in M365:

- **Quarantine Rule** - Quarantines external emails that bypass ETP
- **Reject Rule** - Alternative to quarantine, rejects bypassed emails
- **Allow Rule** - Lets ETP-processed mail through (auto-enabled)
- **SafeLinks Bypass** - Allows ETP quarantine reports to skip SafeLinks

## Quick Start

```powershell
.\Create_OpenText_Email_Threat_Protection_Rules.ps1
```

## After Running

1. Visit [Exchange Admin Center](https://admin.exchange.microsoft.com/#/transportrules)
2. Enable either Quarantine or Reject (not both)
3. Enable SafeLinks Bypass if using ETP
4. Verify rule priorities

## Protected Items

All rules exclude:
- Voicemail messages
- Meeting forwards  
- ETP IP ranges

## Requirements

- PowerShell 5.1+
- Exchange admin permissions

## License

MIT
