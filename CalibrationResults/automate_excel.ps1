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
                    # Check file size
                    $fileSize = (Get-Item $csvFile.FullName).Length
                    $fileSizeMB = [math]::Round($fileSize / 1MB, 2)
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
        
        # Save the processed file
        $outputFileName = "TVN-4-$TVNNumber-$RFGName.xlsx"
        $outputPath = Join-Path $RFGPath $outputFileName
        Write-Host "  Saving: $outputFileName" -ForegroundColor Yellow
        
        try {
            $workbook.SaveAs($outputPath)
            Write-Host "  Saved: $outputFileName" -ForegroundColor Green
        }
        catch {
            Write-Host "  WARNING: Failed to save $outputFileName`: $($_.Exception.Message)" -ForegroundColor Yellow
        }
        
        # Export CSV to main output folder
        Write-Host "  Exporting CSV..." -ForegroundColor Yellow
        try {
            $exportWorksheet = $workbook.Worksheets.Item("Export.CSV")
            $exportWorksheet.Activate()
            
            $csvOutputPath = Join-Path $OutputPath "calibrate_rfg_$rfgNumber.csv"
            $workbook.SaveAs($csvOutputPath, 6) # 6 = CSV format
            Write-Host "  Exported CSV: calibrate_rfg_$rfgNumber.csv" -ForegroundColor Green
        }
        catch {
            Write-Host "  WARNING: Failed to export CSV: $($_.Exception.Message)" -ForegroundColor Yellow
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