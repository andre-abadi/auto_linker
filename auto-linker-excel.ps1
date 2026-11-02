<#
This script:
- Prompts for an Excel file
- Uses the Evidence folder, or Documents if Evidence is not present
- Detects Document ID column in first 3 rows of each worksheet
- Creates hyperlinks between Document IDs and their files
- Displays progress for long-running operations
- Generates separate errata files for unreferenced and missing documents
- Reports total execution time and hyperlinks created
#>

#
# Start Global Variables
#
# maximum number of UsedRange rows to scan for the header
$header_search_rows = 5
# write-host for every x links created
$callout_interval = 50 
# possible column headers to look for
$hyperlink_column_headers = @("Document ID", "Document Identifier")
# set true to show Excel while running
$excel_visible = $false
# set true to generate a transcript log file
$enable_log_file = $false
#
# End Global Variables
#

$scriptStartTime = Get-Date
if ($enable_log_file) {
    Start-Transcript -Path ("auto_linker_log_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".txt") | Out-Null
}

# Select an Excel file from the current directory.
$CurrentDir = (Get-Location).Path
$ExcelFiles = @(Get-ChildItem -LiteralPath $CurrentDir -Filter "*.xlsx" -File | Sort-Object Name)

if ($ExcelFiles.Count -eq 0) {
    Write-Host "No Excel files found in the current directory. Exiting."
    if ($enable_log_file) { Stop-Transcript | Out-Null }
    return
}

Write-Host "`nExcel files in the current directory:"
for ($Index = 0; $Index -lt $ExcelFiles.Count; $Index++) {
    Write-Host "[$($Index + 1)] $($ExcelFiles[$Index].Name)"
}

while ($true) {
    $Selection = Read-Host "Select a workbook by number (blank or 0 to exit)"
    if ([string]::IsNullOrWhiteSpace($Selection) -or $Selection.Trim() -eq "0") {
        Write-Host "No Excel file selected. Exiting."
        if ($enable_log_file) { Stop-Transcript | Out-Null }
        return
    }

    [int]$SelectionNumber = 0
    if ([int]::TryParse($Selection.Trim(), [ref]$SelectionNumber) -and
        $SelectionNumber -ge 1 -and $SelectionNumber -le $ExcelFiles.Count) {
        break
    }

    Write-Host "Enter a number from 1 to $($ExcelFiles.Count), or blank/0 to exit."
}

$ExcelFilePath = $ExcelFiles[$SelectionNumber - 1].FullName

# Prefer Evidence; use Documents only when Evidence is absent.
$DocFolder = $null
foreach ($FolderName in @("Evidence", "Documents")) {
    $CandidateFolder = Join-Path $CurrentDir $FolderName
    if (Test-Path -LiteralPath $CandidateFolder -PathType Container) {
        $DocFolder = (Get-Item -LiteralPath $CandidateFolder).FullName
        break
    }
}

if (-not $DocFolder) {
    Write-Error "Exiting because neither an 'Evidence' nor a 'Documents' folder was found beside the selected Excel file."
    exit 1
}

$enumStart = Get-Date

# Get all files in the document folder
Write-Host "`nEnumerating files in '$(Split-Path $DocFolder -Leaf)'."
$Files = Get-ChildItem -Path $DocFolder -File

$TotalFiles = $Files.Count

# Create a hashtable for quick lookup of files by identifier
$FileLookup = @{}
$DuplicateFileLookup = @{} # Create duplicate file listing
foreach ($File in $Files) {
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
    # Build the matching pattern from the editable header list.
    $HyperlinkColumnHeaderPattern = "^\s*(?:" + (($hyperlink_column_headers | ForEach-Object { [regex]::Escape($_) }) -join "|") + ")\b"

    # Get the count of worksheets
    $WorksheetCount = $Workbook.Worksheets.Count

    # Iterate through each worksheet using a for loop
    for ($w = 1; $w -le $WorksheetCount; $w++) {
        $Worksheet = $Workbook.Worksheets.Item($w)
        # Check if worksheet is hidden (0 = hidden, 2 = very hidden)
        if ($Worksheet.Visible -ne -1) {
            Write-Host "    Skipping hidden worksheet: '$($Worksheet.Name)'"
            # Release $Worksheet
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Worksheet) | Out-Null
            $Worksheet = $null
            continue
        }
        #Write-Host "    Processing worksheet: '$($Worksheet.Name)'"
        Write-Host "    " -NoNewline
        Write-Host $Worksheet.Name -ForegroundColor Magenta -NoNewline
        Write-Host " worksheet found."
        # Find the identifier column in the configured number of rows
        $DocIDColumn = $null
        $HeaderRow = $null  # Track which row contains the header
        $FoundHeader = $false

        for ($Row = 1; $Row -le $header_search_rows; $Row++) {
            $UsedRange = $Worksheet.UsedRange
            $Rows = $UsedRange.Rows
            if ($Row -gt $Rows.Count) {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Rows) | Out-Null
                $Rows = $null
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($UsedRange) | Out-Null
                $UsedRange = $null
                break
            }
            $RowRange = $Rows.Item($Row)
            $Columns = $RowRange.Columns

            $ColumnCount = $Columns.Count
            for ($Col = 1; $Col -le $ColumnCount; $Col++) {
                $Cell = $Columns.Item($Col)
                if ($Cell.Text -match $HyperlinkColumnHeaderPattern) {
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
            Write-Host "        No identifier column found in worksheet '$($Worksheet.Name)'. Skipping."
            # Release $Worksheet
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Worksheet) | Out-Null
            $Worksheet = $null
            continue
        }
        if ($DocIDColumn) {
            # Get the actual header text
            $headerCell = $Worksheet.Cells.Item($HeaderRow, $DocIDColumn)
            $actualHeader = $headerCell.Text.Trim()
    
            # Get the first value in the column (first cell after header)
            $firstValueCell = $Worksheet.Cells.Item($HeaderRow + 1, $DocIDColumn)
            $firstValue = $firstValueCell.Text.Trim()
    
            Write-Host "        Identifier header '$actualHeader' found at column $DocIDColumn."
            Write-Host "        First value: $firstValue."
    
            # Clean up COM objects
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($headerCell) | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($firstValueCell) | Out-Null
            $headerCell = $null
            $firstValueCell = $null
        }
        #Write-Host "    Found a configured document ID header at column $DocIDColumn."
        Write-Host "        Starting hyperlinking."

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
                        Write-Host "        $LinkCounter hyperlinks created."
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


        Write-Host "        Finished hyperlinking."
        Write-Host "        " -NoNewline
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


$NoReferencePath = Join-Path $CurrentDir "_no_reference.txt"
$NoFilesPath = Join-Path $CurrentDir "_no_files.txt"

if ($DuplicateFileLookup.Count -gt 0) {
    $UnreferencedFiles = @($DuplicateFileLookup.Values | Sort-Object)
    Set-Content -LiteralPath $NoReferencePath -Value $UnreferencedFiles
    Write-Host "Unreferenced source files saved to $(Split-Path $NoReferencePath -Leaf) " -NoNewline
    Write-Host "($($UnreferencedFiles.Count))" -ForegroundColor Yellow
} else {
    Write-Host "No unreferenced source files found."
}

if ($MissingFiles.Count -gt 0) {
    $MissingFiles | Sort-Object | Set-Content -LiteralPath $NoFilesPath
    Write-Host "Missing Excel document IDs saved to $(Split-Path $NoFilesPath -Leaf) " -NoNewline
    Write-Host "($($MissingFiles.Count))" -ForegroundColor Yellow
} else {
    Write-Host "No missing source files found."
}



$scriptDuration = (Get-Date) - $scriptStartTime
Write-Host "Total execution time: $($scriptDuration.ToString('hh\:mm\:ss'))`n"

# Stop logging
if ($enable_log_file) { Stop-Transcript | Out-Null }

# Pause before exit
Write-Host "Press ENTER to exit."
$null = Read-Host
