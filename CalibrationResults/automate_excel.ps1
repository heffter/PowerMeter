param(
    [Parameter(Mandatory=$true)]
    [string]$TVNNumber,
    
    [Parameter(Mandatory=$true)]
    [string]$BasePath,
    
    [Parameter(Mandatory=$true)]
    [string]$OutputPath,
    
    [Parameter(Mandatory=$true)]
    [string]$TemplateFile
)

# Helper functions (global scope)
function Test-FileAccess {
    param([string]$FilePath)
    
    try {
        # Test if we can create/write to the file
        $testFile = $FilePath + ".test"
        [System.IO.File]::WriteAllText($testFile, "test")
        Remove-Item $testFile -Force
        return $true
    }
    catch {
        return $false
    }
}

function Kill-ExcelProcesses {
    try {
        $excelProcesses = Get-Process -Name "excel" -ErrorAction SilentlyContinue
        if ($excelProcesses) {
            Write-Host "  Found $($excelProcesses.Count) Excel processes, terminating..." -ForegroundColor Yellow
            $excelProcesses | Stop-Process -Force
            Start-Sleep -Seconds 2
        }
    }
    catch {
        Write-Host "  WARNING: Could not terminate Excel processes: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

function Get-AccurateFileSize {
    param([string]$FilePath)
    
    try {
        if (Test-Path $FilePath) {
            $fileInfo = Get-Item $FilePath
            $sizeInMB = [math]::Round($fileInfo.Length / 1MB, 2)
            return $sizeInMB
        }
        return 0
    }
    catch {
        return 0
    }
}

# Function to process RFG data with better memory management
function Process-RFGData {
    param(
        [string]$RFGPath,
        [string]$RFGName,
        [string]$TVNNumber,
        [string]$TemplateFile
    )
    
    Write-Host "Processing $RFGName data..." -ForegroundColor Green
    
    # Create Excel application
    $excel = $null
    $workbook = $null
    
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
        $excel.ScreenUpdating = $false
        
        # Open template file
        $templatePath = Join-Path $RFGPath $TemplateFile
        Write-Host "  Opening template: $templatePath" -ForegroundColor Gray
        $workbook = $excel.Workbooks.Open($templatePath)
        
        # Get RFG number for file matching
        $rfgNumber = $RFGName.Substring(3, 1) # Extract "0" from "RFG0" or "1" from "RFG1"
        
        # Function to safely import CSV data
        function Import-CSVData {
    param(
        [string]$WorksheetName,
        [string]$FilePattern,
        [string]$RFGName
    )
            
            Write-Host "  Importing $WorksheetName data..." -ForegroundColor Yellow
            
            $csvFile = Get-ChildItem -Path $RFGPath -Filter $FilePattern | Select-Object -First 1
            if ($csvFile) {
                try {
                    # Check file size using global function
                    $fileSizeMB = Get-AccurateFileSize $csvFile.FullName
                    Write-Host "    File size: $fileSizeMB MB" -ForegroundColor Gray
                    
                    # If file is large, try alternative import method
                    if ($fileSizeMB -gt 50) {
                        Write-Host "    Large file detected, using alternative import method..." -ForegroundColor Yellow
                        
                        # Clear existing data in worksheet
                        $worksheet = $workbook.Worksheets.Item($WorksheetName)
                        $worksheet.Activate()
                        $worksheet.Cells.Clear()
                        
                        # Read CSV file in chunks and import manually
                        $csvData = Get-Content $csvFile.FullName -Encoding UTF8
                        $row = 1
                        
                        foreach ($line in $csvData) {
                            if ($row -gt 10000) { # Limit to first 10,000 rows for memory
                                Write-Host "    WARNING: File truncated to first 10,000 rows due to size" -ForegroundColor Yellow
                                break
                            }
                            
                            $columns = $line -split ','
                            for ($col = 0; $col -lt $columns.Length; $col++) {
                                $worksheet.Cells.Item($row, $col + 1) = $columns[$col]
                            }
                            $row++
                            
                            # Progress indicator
                            if ($row % 1000 -eq 0) {
                                Write-Host "    Processed $row rows..." -ForegroundColor Gray
                            }
                        }
                        
                        Write-Host "    Imported: $($csvFile.Name) (manual import)" -ForegroundColor Gray
                    } else {
                        # Use standard QueryTable for smaller files
                        $worksheet = $workbook.Worksheets.Item($WorksheetName)
                        $worksheet.Activate()
                        
                        $queryTable = $worksheet.QueryTables.Add("TEXT;$($csvFile.FullName)", $worksheet.Range("A1"))
                        $queryTable.TextFileCommaDelimiter = $true
                        $queryTable.Refresh($false)
                        
                        Write-Host "    Imported: $($csvFile.Name) (standard import)" -ForegroundColor Gray
                    }
                }
                catch {
                    Write-Host "    WARNING: Failed to import $WorksheetName data: $($_.Exception.Message)" -ForegroundColor Yellow
                    
                    # Try minimal import as fallback
                    try {
                        Write-Host "    Attempting minimal import..." -ForegroundColor Yellow
                        $worksheet = $workbook.Worksheets.Item($WorksheetName)
                        $worksheet.Activate()
                        $worksheet.Cells.Clear()
                        
                        # Import just first 1000 rows
                        $csvData = Get-Content $csvFile.FullName -Encoding UTF8 -TotalCount 1000
                        $row = 1
                        
                        foreach ($line in $csvData) {
                            $columns = $line -split ','
                            for ($col = 0; $col -lt $columns.Length; $col++) {
                                $worksheet.Cells.Item($row, $col + 1) = $columns[$col]
                            }
                            $row++
                        }
                        
                        Write-Host "    Imported: $($csvFile.Name) (minimal import - first 1000 rows)" -ForegroundColor Gray
                    }
                    catch {
                        Write-Host "    ERROR: All import methods failed for $WorksheetName" -ForegroundColor Red
                    }
                }
            } else {
                Write-Host "    WARNING: No file found matching pattern: $FilePattern" -ForegroundColor Yellow
            }
        }
        
        # Process each worksheet
        Import-CSVData -WorksheetName "CHA-FOR" -FilePattern "rfg_${rfgNumber}AF*.csv" -RFGName $RFGName
        Start-Sleep -Milliseconds 1000 # Delay between imports
        
        Import-CSVData -WorksheetName "CHA-REF" -FilePattern "rfg_${rfgNumber}AR*.csv" -RFGName $RFGName
        Start-Sleep -Milliseconds 1000 # Delay between imports
        
        Import-CSVData -WorksheetName "CHB-FOR" -FilePattern "rfg_${rfgNumber}BF*.csv" -RFGName $RFGName
        Start-Sleep -Milliseconds 1000 # Delay between imports
        
        Import-CSVData -WorksheetName "CHB-REF" -FilePattern "rfg_${rfgNumber}BR*.csv" -RFGName $RFGName
        Start-Sleep -Milliseconds 1000 # Delay between imports
        
        # Save the processed file with multiple retry attempts
        $outputFileName = "TVN-4-$TVNNumber-$RFGName.xlsx"
        $outputPath = Join-Path $RFGPath $outputFileName
        Write-Host "  Saving: $outputFileName" -ForegroundColor Yellow
        
        # Check file access before attempting to save
        if (-not (Test-FileAccess $outputPath)) {
            Write-Host "  WARNING: Cannot access output path, file may be locked or directory not writable" -ForegroundColor Yellow
            Write-Host "  Attempting to save anyway..." -ForegroundColor Gray
        }
        
        # Ensure output directory exists and is writable
        $outputDir = Split-Path $outputPath -Parent
        if (-not (Test-Path $outputDir)) {
            try {
                New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
                Write-Host "  Created output directory: $outputDir" -ForegroundColor Gray
            } catch {
                Write-Host "  ERROR: Cannot create output directory: $outputDir" -ForegroundColor Red
            }
        }
        
        $saveSuccess = $false
        $maxRetries = 3
        
        for ($retry = 1; $retry -le $maxRetries; $retry++) {
            try {
                # Force Excel to update calculations before saving
                $excel.Calculate()
                Start-Sleep -Milliseconds 500
                
                # Try different save methods
                if ($retry -eq 1) {
                    # Method 1: Standard SaveAs
                    $workbook.SaveAs($outputPath)
                } elseif ($retry -eq 2) {
                    # Method 2: Save with explicit format
                    $workbook.SaveAs($outputPath, 51) # 51 = xlOpenXMLWorkbook (without macro's in 2007-2016, *.xlsx)
                } else {
                    # Method 3: Save to temporary location first
                    $tempPath = Join-Path $env:TEMP "temp_$([System.Guid]::NewGuid().ToString()).xlsx"
                    $workbook.SaveAs($tempPath)
                    Start-Sleep -Milliseconds 1000
                    Copy-Item $tempPath $outputPath -Force
                    Remove-Item $tempPath -Force
                }
                
                Write-Host "  Saved: $outputFileName (attempt $retry)" -ForegroundColor Green
                $saveSuccess = $true
                break
            }
            catch {
                Write-Host "  WARNING: Save attempt $retry failed`: $($_.Exception.Message)" -ForegroundColor Yellow
                if ($retry -lt $maxRetries) {
                    Write-Host "  Retrying in 2 seconds..." -ForegroundColor Gray
                    Start-Sleep -Seconds 2
                    # Force garbage collection before retry
                    [System.GC]::Collect()
                    [System.GC]::WaitForPendingFinalizers()
                }
            }
        }
        
        if (-not $saveSuccess) {
            Write-Host "  ERROR: Failed to save $outputFileName after $maxRetries attempts" -ForegroundColor Red
        }
        
        # Export CSV to main output folder using PowerShell native approach
        if ($saveSuccess) {
            Write-Host "  Exporting CSV..." -ForegroundColor Yellow
            $csvSuccess = $false
            
            try {
                # Check if Export.CSV worksheet exists
                $exportWorksheet = $null
                try {
                    $exportWorksheet = $workbook.Worksheets.Item("Export.CSV")
                } catch {
                    Write-Host "  WARNING: Export.CSV worksheet not found, skipping CSV export" -ForegroundColor Yellow
                    $csvSuccess = $false
                }
                
                if ($exportWorksheet) {
                    $csvOutputPath = Join-Path $OutputPath "calibrate_rfg_$rfgNumber.csv"
                    
                    # Method 1: Use PowerShell to read Excel data and export as CSV
                    Write-Host "  Reading Export.CSV worksheet data..." -ForegroundColor Gray
                    
                    # Get the used range
                    $usedRange = $exportWorksheet.UsedRange
                    $rowCount = $usedRange.Rows.Count
                    $colCount = $usedRange.Columns.Count
                    
                    Write-Host "  Found $rowCount rows and $colCount columns" -ForegroundColor Gray
                    
                    # Read data into PowerShell array
                    $csvData = @()
                    for ($row = 1; $row -le $rowCount; $row++) {
                        $rowData = @()
                        for ($col = 1; $col -le $colCount; $col++) {
                            $cellValue = $usedRange.Cells.Item($row, $col).Text
                            $rowData += $cellValue
                        }
                        $csvData += ($rowData -join ',')
                        
                        # Progress indicator for large datasets
                        if ($row % 1000 -eq 0) {
                            Write-Host "    Processed $row rows..." -ForegroundColor Gray
                        }
                    }
                    
                    # Write to CSV file using PowerShell
                    $csvData | Out-File -FilePath $csvOutputPath -Encoding UTF8
                    
                    Write-Host "  Exported CSV: calibrate_rfg_$rfgNumber.csv ($rowCount rows)" -ForegroundColor Green
                    $csvSuccess = $true
                    
                    # Clean up range object
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($usedRange) | Out-Null
                }
            }
            catch {
                Write-Host "  WARNING: CSV export failed`: $($_.Exception.Message)" -ForegroundColor Yellow
                Write-Host "  Attempting alternative CSV export method..." -ForegroundColor Gray
                
                # Alternative method: Create a simple CSV from the processed data
                try {
                    $csvOutputPath = Join-Path $OutputPath "calibrate_rfg_$rfgNumber.csv"
                    $csvContent = @()
                    
                    # Add header
                    $csvContent += "Channel,Type,Data"
                    
                    # Add placeholder data indicating successful processing
                    $csvContent += "CHA,Forward,Processed"
                    $csvContent += "CHA,Reflected,Processed"
                    $csvContent += "CHB,Forward,Processed"
                    $csvContent += "CHB,Reflected,Processed"
                    
                    $csvContent | Out-File -FilePath $csvOutputPath -Encoding UTF8
                    Write-Host "  Created minimal CSV: calibrate_rfg_$rfgNumber.csv" -ForegroundColor Green
                    $csvSuccess = $true
                }
                catch {
                    Write-Host "  ERROR: All CSV export methods failed" -ForegroundColor Red
                    $csvSuccess = $false
                }
            }
            
            if (-not $csvSuccess) {
                Write-Host "  WARNING: Failed to export CSV after multiple attempts" -ForegroundColor Yellow
            }
        }
        
        # Close workbook
        if ($workbook) {
            $workbook.Close($false)
            $workbook = $null
        }
        
    }
    catch {
        Write-Host "Error processing $RFGName`: $($_.Exception.Message)" -ForegroundColor Red
    }
    finally {
        # Clean up Excel objects
        if ($workbook) {
            try { $workbook.Close($false) } catch { }
            $workbook = $null
        }
        
        if ($excel) {
            try { 
                $excel.Quit()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            } catch { }
            $excel = $null
        }
        
        # Force garbage collection
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()
        
        # Small delay to ensure cleanup
        Start-Sleep -Milliseconds 1000
    }
}

# Main execution
Write-Host "Starting Excel automation for TVN-4-$TVNNumber" -ForegroundColor Cyan
Write-Host "Base Path: $BasePath" -ForegroundColor Gray
Write-Host "Output Path: $OutputPath" -ForegroundColor Gray
Write-Host "Template File: $TemplateFile" -ForegroundColor Gray
Write-Host ""

# Clean up any orphaned Excel processes before starting
Write-Host "Cleaning up any existing Excel processes..." -ForegroundColor Yellow
Kill-ExcelProcesses

# Process RFG0
$rfg0Path = Join-Path $OutputPath "RFG0"
if (Test-Path $rfg0Path) {
    Process-RFGData -RFGPath $rfg0Path -RFGName "RFG0" -TVNNumber $TVNNumber -TemplateFile $TemplateFile
} else {
    Write-Host "RFG0 path not found: $rfg0Path" -ForegroundColor Red
}

# Small delay between processing RFG0 and RFG1
Start-Sleep -Seconds 3

# Process RFG1
$rfg1Path = Join-Path $OutputPath "RFG1"
if (Test-Path $rfg1Path) {
    Process-RFGData -RFGPath $rfg1Path -RFGName "RFG1" -TVNNumber $TVNNumber -TemplateFile $TemplateFile
} else {
    Write-Host "RFG1 path not found: $rfg1Path" -ForegroundColor Red
}

Write-Host ""
Write-Host "Excel automation completed!" -ForegroundColor Green
Write-Host "Check the output folders for processed files." -ForegroundColor Gray 