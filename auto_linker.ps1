<#
This script:
- Processes a single Excel file in the current directory
- Searches for a single folder containing corresponding files
- Detects Document ID column in first 3 rows of each worksheet
- Creates hyperlinks between Document IDs and their files
- Displays progress for long-running operations
- Generates errata file listing missing/extra documents
- Reports total execution time and hyperlinks created
#>

<#
Changes:
- V11
  - Folder agnostic, no name requirement just that there's only one.
  - Document ID columns can have words after Document ID and still be included.
  - Declare column header in case of changing away from Document ID.
  - Add transcript logging.
#>

#
# Start Global Variables
#
# write-host for every x files enumerated and links created
$callout_interval = 5000 
# string to look for to link
$hyperlink_column_header = "Document ID"
# set true to show Excel while running
$excel_visible = $true
#
# End Global Variables
#

$scriptStartTime = Get-Date
Start-Transcript -Path ("auto_linker_log_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".log")  | Out-Null

# Load the necessary assemblies for Excel automation
Add-Type -AssemblyName Microsoft.Office.Interop.Excel

# Set the current directory path as a string
$CurrentDir = (Get-Location).Path

# Find the Excel file in the current directory
$ExcelFiles = Get-ChildItem -Path $CurrentDir -Filter *.xlsx

if ($ExcelFiles.Count -ne 1) {
    Write-Error "Exiting because there should be exactly one Excel file in the current directory and this condition was not met."
    exit 1
}

$ExcelFilePath = $ExcelFiles[0].FullName

# Find the single subfolder in the current directory
$Folders = Get-ChildItem -Path $CurrentDir -Directory
if ($Folders.Count -ne 1) {
    Write-Error "Exiting because there should be exactly one folder in the current directory and this condition was not met."
    exit 1
}
$DocFolder = $Folders[0].FullName

$enumStart = Get-Date

# Get all files in the document folder
Write-Host "`nEnumerating files in '$(Split-Path $DocFolder -Leaf)'."
$Files = Get-ChildItem -Path $DocFolder -File

$TotalFiles = $Files.Count
$FileCounter = 0

# Create a hashtable for quick lookup of files by identifier
$FileLookup = @{}
$DuplicateFileLookup = @{} # Create duplicate file listing
foreach ($File in $Files) {
    $FileCounter++
    if ($FileCounter % $callout_interval -eq 0) {
        Write-Host "    $FileCounter files enumerated."
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }

    # Extract the filename without extension
    $DocID = $File.BaseName.Trim().ToUpper()
    if (-not $FileLookup.ContainsKey($DocID)) {
        $FileLookup[$DocID] = $File.Name
        $DuplicateFileLookup[$DocID] = $File.Name # Populate duplicate structure
    } else {
        Write-Error "Exiting because multiple files found with the identifier: $DocID"
        exit 1
    }
}
Write-Host "    File enumeration completed."
$enumDuration = (Get-Date) - $enumStart
Write-Host "    File enumeration took: $($enumDuration.ToString('hh\:mm\:ss'))"
Write-Host "    " -NoNewline
Write-Host $TotalFiles -ForegroundColor Green -NoNewline
Write-Host " referrable files found."



# Prepare hashtables for missing and extra files
$LinkedDocIDs = @{}
$MissingFiles = New-Object System.Collections.Generic.HashSet[string]
$TotalHyperlinksAdded = 0  # Initialize total hyperlinks counter

# Open Excel application
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $excel_visible
$Excel.DisplayAlerts = $false
$Excel.ScreenUpdating = $false
$Excel.EnableAnimations = $false

# Open the Excel workbook
Write-Host "`nOpening Excel File"
$Workbook = $Excel.Workbooks.Open($ExcelFilePath)

try {
    # Get the count of worksheets
    $WorksheetCount = $Workbook.Worksheets.Count

    # Iterate through each worksheet using a for loop
    for ($w = 1; $w -le $WorksheetCount; $w++) {
        $Worksheet = $Workbook.Worksheets.Item($w)
        #Write-Host "    Processing worksheet: '$($Worksheet.Name)'"
        Write-Host "    " -NoNewline
        Write-Host $Worksheet.Name -ForegroundColor Magenta -NoNewline
        Write-Host " worksheet found."
        # Find the identifier column in the first three rows
        $DocIDColumn = $null
        $HeaderRow = $null  # Track which row contains the header
        $FoundHeader = $false

        for ($Row = 1; $Row -le 3; $Row++) {
            $UsedRange = $Worksheet.UsedRange
            $Rows = $UsedRange.Rows
            $RowRange = $Rows.Item($Row)
            $Columns = $RowRange.Columns

            $ColumnCount = $Columns.Count
            for ($Col = 1; $Col -le $ColumnCount; $Col++) {
                $Cell = $Columns.Item($Col)
                if ($Cell.Text -match "^\s*$DOCUMENT_ID_HEADER\b") {
                    $DocIDColumn = $Cell.Column
                    $HeaderRow = $Row    # Store which row contains the header
                    $FoundHeader = $true
                    # Release $Cell
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Cell) | Out-Null
                    $Cell = $null
                    break
                }
                # Release $Cell
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Cell) | Out-Null
                $Cell = $null
            }

            # Release COM objects in the loop
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Columns) | Out-Null
            $Columns = $null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($RowRange) | Out-Null
            $RowRange = $null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Rows) | Out-Null
            $Rows = $null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($UsedRange) | Out-Null
            $UsedRange = $null

            if ($FoundHeader) { break }
        }

        if (-not $DocIDColumn) {
            Write-Host "    No identifier column found in worksheet '$($Worksheet.Name)'. Skipping."
            # Release $Worksheet
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Worksheet) | Out-Null
            $Worksheet = $null
            continue
        }
        Write-Host "    Found identifier column $DocIDColumn."
        Write-Host "    Starting hyperlinking."

        # Get the used range starting from the row after the header
        $UsedRange = $Worksheet.UsedRange
        $StartRow = $UsedRange.Row + $HeaderRow  # Start from the row after where we found the header
        $EndRow = $UsedRange.Row + $UsedRange.Rows.Count - 1
        $LinkCounter = 0

        for ($Row = $StartRow; $Row -le $EndRow; $Row++) {
            $Cell = $Worksheet.Cells.Item($Row, $DocIDColumn)
            $DocID = $Cell.Text.Trim().ToUpper()
            if ($DocID) {
                if ($FileLookup.ContainsKey($DocID)) {
                    $RelativePath = Join-Path (Split-Path $DocFolder -Leaf) $FileLookup[$DocID]
                    # Add hyperlink to the cell
                    $null = $Worksheet.Hyperlinks.Add($Cell, $RelativePath)
                    $LinkCounter++
                    $TotalHyperlinksAdded++
                    $LinkedDocIDs[$DocID] = $true
                    $DuplicateFileLookup.Remove($DocID) # Remove file from duplicate structure
                    if ($LinkCounter % $callout_interval -eq 0) {
                        Write-Host "    $LinkCounter hyperlinks created."
                        # garbage collection
                        [System.GC]::Collect()
                        [System.GC]::WaitForPendingFinalizers()
                    }
                } else {
                    # Identifier found but no corresponding file
                    $MissingFiles.Add($DocID) | Out-Null
                }
            }
            # Release $Cell
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Cell) | Out-Null
            $Cell = $null
        }


        Write-Host "    Finished hyperlinking."
        Write-Host "    " -NoNewline
        Write-Host $LinkCounter -ForegroundColor Green -NoNewline
        Write-Host " hyperlinks created."

        # Release COM objects after worksheet processing
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($UsedRange) | Out-Null
        $UsedRange = $null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Worksheet) | Out-Null
        $Worksheet = $null

        # Force garbage collection
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}
catch {
    Write-Error "An error occurred: $_"
}
finally {
    # Close the workbook (save changes)
    $Workbook.Save()
    $Workbook.Close($false)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook) | Out-Null
    $Workbook = $null

    $Excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) | Out-Null
    $Excel = $null

    # Force garbage collection
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Host "Workbook closed.`n"
}

# Add line break and total hyperlinks summary before errata handling
Write-Host "Total hyperlinks created across all worksheets: " -NoNewline
Write-Host $TotalHyperlinksAdded -ForegroundColor Green


if ($DuplicateFileLookup.Keys.Count -gt 0 -or $MissingFiles.Count -gt 0) {
    # Create errata file with timestamp
    $ErrataFilePath = Join-Path $CurrentDir ("errata_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".txt")
    
    # Output extraneous files if they exist
    if ($DuplicateFileLookup.Keys.Count -gt 0) {
        Add-Content -Path $ErrataFilePath -Value "Extraneous Files:"
        foreach ($ExtraFile in $DuplicateFileLookup.Keys) {
            Add-Content -Path $ErrataFilePath -Value $DuplicateFileLookup[$ExtraFile]
        }
    }
    
    # Output missing files if they exist
    if ($MissingFiles.Count -gt 0) {
        Add-Content -Path $ErrataFilePath -Value "`nMissing Files:"
        foreach ($MissingFile in $MissingFiles) {
            Add-Content -Path $ErrataFilePath -Value $MissingFile
        }
    }
    
    # Output confirmation message to console
    Write-Host "Errata saved to $ErrataFilePath"
} else {
    Write-Host "No errata file needed - all identifiers were processed successfully!"
}



$scriptDuration = (Get-Date) - $scriptStartTime
Write-Host "Total execution time: $($scriptDuration.ToString('hh\:mm\:ss'))`n"

# Stop logging
Stop-Transcript | Out-Null

# Pause before exit
Write-Host "Press ENTER to exit."
$null = Read-Host
