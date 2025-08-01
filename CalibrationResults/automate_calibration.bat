@echo off
setlocal enabledelayedexpansion

echo PowerMeter Calibration Data Automation
echo ======================================
echo.

:: Get TVN system number from user
set /p TVN_NUMBER="Enter TVN system number (e.g., 0106): "

:: Get current date in YYYY-MM-DD format
for /f "tokens=2 delims==" %%a in ('wmic OS Get localdatetime /value') do set "dt=%%a"
set "YEAR=%dt:~0,4%"
set "MONTH=%dt:~4,2%"
set "DAY=%dt:~6,2%"
set "DATE_FOLDER=%YEAR%-%MONTH%-%DAY%"

:: Set base paths
set BASE_PATH=C:\AMS\Remote_Control\log
set OUTPUT_PATH=%BASE_PATH%\TVN-4-%TVN_NUMBER%-%DATE_FOLDER%
set RFG0_PATH=%OUTPUT_PATH%\RFG0
set RFG1_PATH=%OUTPUT_PATH%\RFG1

echo.
echo Step 1: Renaming log files to CSV...
if exist "%BASE_PATH%\rfg_*.log" (
    for %%f in ("%BASE_PATH%\rfg_*.log") do (
        ren "%%f" "%%~nf.csv"
        echo Renamed: %%~nxf to %%~nf.csv
    )
) else (
    echo No log files found to rename.
)

echo.
echo Step 2: Creating folder structure...
if not exist "%RFG0_PATH%" mkdir "%RFG0_PATH%"
if not exist "%RFG1_PATH%" mkdir "%RFG1_PATH%"

echo Step 3: Finding and copying template files...
:: Find any .xltx file in the base path
for %%f in ("%BASE_PATH%\*.xltx") do (
    set TEMPLATE_FILE=%%~nxf
    echo Found template: !TEMPLATE_FILE!
    copy "%%f" "%RFG0_PATH%\"
    copy "%%f" "%RFG1_PATH%\"
    echo Templates copied successfully.
    goto :template_found
)

:: If no .xltx found, try .xlsx
for %%f in ("%BASE_PATH%\*.xlsx") do (
    set TEMPLATE_FILE=%%~nxf
    echo Found template: !TEMPLATE_FILE!
    copy "%%f" "%RFG0_PATH%\"
    copy "%%f" "%RFG1_PATH%\"
    echo Templates copied successfully.
    goto :template_found
)

echo ERROR: No template file (.xltx or .xlsx) found in %BASE_PATH%
pause
exit /b 1

:template_found

echo Step 4: Copying and organizing CSV files...
echo Processing RFG files...

:: Copy RFG0 files (rfg_0*.csv)
for %%f in ("%BASE_PATH%\rfg_0*.csv") do (
    echo Copying %%f to RFG0 folder...
    copy "%%f" "%RFG0_PATH%\"
)

:: Copy RFG1 files (rfg_1*.csv) 
for %%f in ("%BASE_PATH%\rfg_1*.csv") do (
    echo Copying %%f to RFG1 folder...
    copy "%%f" "%RFG1_PATH%\"
)

echo.
echo Step 5: Running Excel automation...
echo This will open Excel and process the data automatically.
echo Please wait for Excel to complete the processing...

:: Run PowerShell script for Excel automation
powershell -ExecutionPolicy Bypass -File "%~dp0automate_excel.ps1" -TVNNumber "%TVN_NUMBER%" -BasePath "%BASE_PATH%" -OutputPath "%OUTPUT_PATH%" -TemplateFile "!TEMPLATE_FILE!"

echo.
echo Automation completed!
echo Check the following folders for results:
echo - %RFG0_PATH%\ (RFG0 calibration data)
echo - %RFG1_PATH%\ (RFG1 calibration data)
echo - %OUTPUT_PATH%\ (Exported CSV files)
echo.
pause 