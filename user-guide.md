# Defender Hunting Solution
## User Guide

### Introduction
The Defender Hunting Solution is a powerful tool designed to automate threat hunting activities using Microsoft Defender Advanced Hunting. This solution enables security analysts to efficiently process Indicators of Compromise (IoCs) and search for potential threats across their environment.

### Prerequisites
Before using the solution, ensure you have:
- PowerShell 5.1 or higher installed
- Access to Microsoft Defender Advanced Hunting
- The following information from your Microsoft Defender environment:
  - Tenant ID
  - Client ID
  - Client Secret

### Installation Steps
1. Create the base directory structure:
```powershell
New-Item -ItemType Directory -Path "C:\DefenderHunting" -Force
New-Item -ItemType Directory -Path "C:\DefenderHunting\Modules\DefenderHunter" -Force
```

2. Save the provided `DefenderHunter.psm1` module file to:
```
C:\DefenderHunting\Modules\DefenderHunter\DefenderHunter.psm1
```

3. Create a configuration file (`config.json`) with your credentials:
```json
{
    "TenantId": "your-tenant-id",
    "ClientId": "your-client-id",
    "ClientSecret": "your-client-secret"
}
```

### Using the Solution
1. **Prepare your IoCs file**:
   - Create a CSV file containing your IoCs
   - Required columns: FamilyName, FileHashes
   - FileHashes should be in JSON format containing MD5, SHA1, and SHA256 values

2. **Import the module**:
```powershell
Import-Module "C:\DefenderHunting\Modules\DefenderHunter\DefenderHunter.psm1" -Force
```

3. **Load configuration**:
```powershell
$config = Get-Content "C:\DefenderHunting\config.json" | ConvertFrom-Json
```

4. **Run the hunting process**:
```powershell
Start-DefenderHunting -IocsFile "path\to\your\iocs.csv" `
    -TenantId $config.TenantId `
    -ClientId $config.ClientId `
    -ClientSecret $config.ClientSecret `
    -LookbackDays 30
```

### Understanding the Output
The solution creates several outputs:

1. **Queries Directory** (`C:\DefenderHunting\Queries`):
   - Contains generated KQL queries for each malware family
   - Files are named using sanitized family names

2. **Exports Directory** (`***REMOVED***`):
   - Contains timestamped folders for each run
   - Each folder includes:
     - CSV files with detection results
     - Summary files for each family
     - A global summary of all detections

3. **Logs Directory** (`C:\DefenderHunting\Logs`):
   - Contains execution logs
   - Useful for troubleshooting

### Common Operations
- **Modify lookback period**:
  ```powershell
  Start-DefenderHunting ... -LookbackDays 60
  ```

- **Enable verbose output**:
  ```powershell
  Start-DefenderHunting ... -Verbose
  ```

### Troubleshooting
If you encounter issues:
1. Check the log files in the Logs directory
2. Verify API credentials in config.json
3. Ensure IoCs CSV file format is correct
4. Check network connectivity to Microsoft Defender API

### Best Practices
- Regularly update your IoCs file
- Keep credentials secure and rotate them periodically
- Review generated queries before running them
- Archive old exports regularly to manage disk space

### Support
For issues or questions:
1. Check the logs for detailed error messages
2. Review the Microsoft Defender Advanced Hunting documentation
3. Contact your security team administrator

This solution streamlines the threat hunting process by automating IoC searches across your environment using Microsoft Defender Advanced Hunting capabilities.
