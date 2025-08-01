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

# Function to process RFG data
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
        
        # Process CHA-FOR (Channel A Forward)
        Write-Host "  Importing CHA-FOR data..." -ForegroundColor Yellow
        $worksheet = $workbook.Worksheets.Item("CHA-FOR")
        $worksheet.Activate()
        
        $csvFile = Get-ChildItem -Path $RFGPath -Filter "rfg_${rfgNumber}AF*.csv" | Select-Object -First 1
        if ($csvFile) {
            try {
                $queryTable = $worksheet.QueryTables.Add("TEXT;$($csvFile.FullName)", $worksheet.Range("A1"))
                $queryTable.TextFileCommaDelimiter = $true
                $queryTable.Refresh($false)
                Write-Host "    Imported: $($csvFile.Name)" -ForegroundColor Gray
            }
            catch {
                Write-Host "    WARNING: Failed to import CHA-FOR data: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        } else {
            Write-Host "    WARNING: No CHA-FOR file found for $RFGName" -ForegroundColor Yellow
        }
        
        # Process CHA-REF (Channel A Reflected)
        Write-Host "  Importing CHA-REF data..." -ForegroundColor Yellow
        $worksheet = $workbook.Worksheets.Item("CHA-REF")
        $worksheet.Activate()
        
        $csvFile = Get-ChildItem -Path $RFGPath -Filter "rfg_${rfgNumber}AR*.csv" | Select-Object -First 1
        if ($csvFile) {
            try {
                $queryTable = $worksheet.QueryTables.Add("TEXT;$($csvFile.FullName)", $worksheet.Range("A1"))
                $queryTable.TextFileCommaDelimiter = $true
                $queryTable.Refresh($false)
                Write-Host "    Imported: $($csvFile.Name)" -ForegroundColor Gray
            }
            catch {
                Write-Host "    WARNING: Failed to import CHA-REF data: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        } else {
            Write-Host "    WARNING: No CHA-REF file found for $RFGName" -ForegroundColor Yellow
        }
        
        # Process CHB-FOR (Channel B Forward)
        Write-Host "  Importing CHB-FOR data..." -ForegroundColor Yellow
        $worksheet = $workbook.Worksheets.Item("CHB-FOR")
        $worksheet.Activate()
        
        $csvFile = Get-ChildItem -Path $RFGPath -Filter "rfg_${rfgNumber}BF*.csv" | Select-Object -First 1
        if ($csvFile) {
            try {
                $queryTable = $worksheet.QueryTables.Add("TEXT;$($csvFile.FullName)", $worksheet.Range("A1"))
                $queryTable.TextFileCommaDelimiter = $true
                $queryTable.Refresh($false)
                Write-Host "    Imported: $($csvFile.Name)" -ForegroundColor Gray
            }
            catch {
                Write-Host "    WARNING: Failed to import CHB-FOR data: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        } else {
            Write-Host "    WARNING: No CHB-FOR file found for $RFGName" -ForegroundColor Yellow
        }
        
        # Process CHB-REF (Channel B Reflected)
        Write-Host "  Importing CHB-REF data..." -ForegroundColor Yellow
        $worksheet = $workbook.Worksheets.Item("CHB-REF")
        $worksheet.Activate()
        
        $csvFile = Get-ChildItem -Path $RFGPath -Filter "rfg_${rfgNumber}BR*.csv" | Select-Object -First 1
        if ($csvFile) {
            try {
                $queryTable = $worksheet.QueryTables.Add("TEXT;$($csvFile.FullName)", $worksheet.Range("A1"))
                $queryTable.TextFileCommaDelimiter = $true
                $queryTable.Refresh($false)
                Write-Host "    Imported: $($csvFile.Name)" -ForegroundColor Gray
            }
            catch {
                Write-Host "    WARNING: Failed to import CHB-REF data: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        } else {
            Write-Host "    WARNING: No CHB-REF file found for $RFGName" -ForegroundColor Yellow
        }
        
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
        Start-Sleep -Milliseconds 500
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
Start-Sleep -Seconds 2

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