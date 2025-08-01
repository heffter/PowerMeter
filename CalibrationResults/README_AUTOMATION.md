# PowerMeter Calibration Data Automation

This automation solution replaces the manual steps described in `ToDo.txt` with automated scripts for Windows 8.1.

## Prerequisites

- Windows 8.1 or later
- Microsoft Excel 2010 or later installed
- PowerShell execution policy allowing script execution
- Log files in the correct location with proper naming convention

## Files Included

1. **`automate_calibration.bat`** - Main batch script for Windows automation
2. **`automate_excel.ps1`** - PowerShell script for Excel automation
3. **`README_AUTOMATION.md`** - This documentation file

## How It Works

The automation performs all the manual steps from the ToDo.txt file:

### Step 1: Log File Renaming
- Automatically renames all `rfg_*.log` files to `rfg_*.csv` format
- Handles the conversion from log to CSV format

### Step 2: File Preparation
- Creates folder structure: `C:\TVN-4-XXXX\RFG0\` and `C:\TVN-4-XXXX\RFG1\`
- Automatically finds and copies any Excel template file (.xltx or .xlsx) to both RFG folders
- Intelligently copies CSV files based on naming convention:
  - Files starting with `rfg_0` → RFG0 folder
  - Files starting with `rfg_1` → RFG1 folder

### Step 3: Excel Data Import
- Opens Excel template files
- Imports CSV data into correct worksheets:
  - CHA-FOR (Channel A Forward)
  - CHA-REF (Channel A Reflected) 
  - CHB-FOR (Channel B Forward)
  - CHB-REF (Channel B Reflected)
- Saves processed files with TVN system number

### Step 4: CSV Export
- Exports processed data as CSV files
- Creates `calibrate_rfg_0.csv` and `calibrate_rfg_1.csv`
- Saves to main output folder `C:\TVN-4-XXXX\`

## Usage Instructions

### Method 1: Simple Execution
1. Double-click `automate_calibration.bat`
2. Enter your TVN system number when prompted (e.g., 0106)
3. Wait for the automation to complete

### Method 2: Command Line
```cmd
cd CalibrationResults
automate_calibration.bat
```

### Method 3: PowerShell Direct
```powershell
powershell -ExecutionPolicy Bypass -File "automate_excel.ps1" -TVNNumber "0106" -BasePath "C:\AMS\Remote_Control\log" -OutputPath "C:\TVN-4-0106" -TemplateFile "RF_Power_Calibrate_Template-0.25-1.2MHz.xltx"
```

## File Structure Expected

### Input Files (in `C:\AMS\Remote_Control\log\`)
- **Template file**: Any `.xltx` or `.xlsx` file (e.g., `RF_Power_Calibrate_Template-0.25-1.2MHz.xltx`)
- `rfg_0AF.RFG_1.Forward.Aug_01_2025-05_25_28.channel_A.Table_0AF_1.log` - RFG0 Channel A Forward data
- `rfg_0AR.RFG_1.Reflected.Aug_01_2025-05_25_28.channel_A.Table_0AR_1.log` - RFG0 Channel A Reflected data  
- `rfg_0BF.RFG_1.Forward.Aug_01_2025-05_25_28.channel_B.Table_0BF_1.log` - RFG0 Channel B Forward data
- `rfg_0BR.RFG_1.Reflected.Aug_01_2025-05_25_28.channel_B.Table_0BR_1.log` - RFG0 Channel B Reflected data
- `rfg_1AF.RFG_2.Forward.Aug_01_2025-05_25_28.channel_A.Table_1AF_1.log` - RFG1 Channel A Forward data
- `rfg_1AR.RFG_2.Reflected.Aug_01_2025-05_25_28.channel_A.Table_1AR_1.log` - RFG1 Channel A Reflected data
- `rfg_1BF.RFG_2.Forward.Aug_01_2025-05_25_28.channel_B.Table_1BF_1.log` - RFG1 Channel B Forward data
- `rfg_1BR.RFG_2.Reflected.Aug_01_2025-05_25_28.channel_B.Table_1BR_1.log` - RFG1 Channel B Reflected data

### Template File Support
The automation automatically detects and uses:
- **Primary**: Any `.xltx` file (Excel template format)
- **Fallback**: Any `.xlsx` file (Excel workbook format)
- **Priority**: First `.xltx` file found, then first `.xlsx` file if no `.xltx` exists

### File Naming Convention
The automation intelligently parses file names:
- **RFG0 files**: Start with `rfg_0` (0AF, 0AR, 0BF, 0BR)
- **RFG1 files**: Start with `rfg_1` (1AF, 1AR, 1BF, 1BR)
- **Channel mapping**:
  - `*AF` = Channel A Forward
  - `*AR` = Channel A Reflected
  - `*BF` = Channel B Forward
  - `*BR` = Channel B Reflected

### Output Structure
```
C:\TVN-4-XXXX\
├── RFG0\
│   ├── [Template File].xltx
│   ├── rfg_0AF*.csv, rfg_0AR*.csv, rfg_0BF*.csv, rfg_0BR*.csv
│   └── TVN-4-XXXX-RFG0.xlsx
├── RFG1\
│   ├── [Template File].xltx
│   ├── rfg_1AF*.csv, rfg_1AR*.csv, rfg_1BF*.csv, rfg_1BR*.csv
│   └── TVN-4-XXXX-RFG1.xlsx
├── calibrate_rfg_0.csv
└── calibrate_rfg_1.csv
```

## Troubleshooting

### Common Issues

1. **PowerShell Execution Policy Error**
   - Run as Administrator: `Set-ExecutionPolicy RemoteSigned`
   - Or use: `powershell -ExecutionPolicy Bypass -File script.ps1`

2. **Template File Not Found**
   - Ensure at least one `.xltx` or `.xlsx` file exists in `C:\AMS\Remote_Control\log\`
   - Check file permissions
   - Verify file extensions are correct

3. **Log Files Not Found**
   - Verify all 8 log files exist in the log folder
   - Check file naming convention (rfg_0AF, rfg_0AR, etc.)
   - Ensure files have `.log` extension

4. **Excel Automation Fails**
   - Ensure Excel is properly installed
   - Close any open Excel instances before running
   - Check Windows COM object permissions

### Error Messages

- **"No template file (.xltx or .xlsx) found"** - Check template file location and extensions
- **"No log files found to rename"** - Check log file location and naming
- **"RFG0 path not found"** - Check folder creation permissions
- **"Error processing RFG0"** - Check Excel installation and COM permissions
- **"WARNING: No CHA-FOR file found"** - Check CSV file naming and location

## Manual Override

If automation fails, you can still perform steps manually:

1. Run only the batch file to create folders and copy files
2. Open Excel templates manually and import CSV data
3. Save and export files as described in original ToDo.txt

## Performance Notes

- Processing time: ~2-5 minutes depending on data size
- Excel runs in background (not visible)
- Progress is shown in console window
- Memory usage: ~50-100MB during processing

## Support

For issues with the automation scripts:
1. Check the console output for error messages
2. Verify all prerequisites are met
3. Ensure file paths and permissions are correct
4. Test with a small dataset first 