# PowerShell v4 Script to Hyperlink Document IDs in Excel with Progress Updates and Errata Logging

# Load the necessary assemblies for Excel automation
Add-Type -AssemblyName Microsoft.Office.Interop.Excel

# Set the current directory path as a string
$CurrentDir = (Get-Location).Path

# Find the Excel file in the current directory
$ExcelFiles = Get-ChildItem -Path $CurrentDir -Filter *.xlsx

if ($ExcelFiles.Count -ne 1) {
    Write-Error "There should be exactly one Excel file in the current directory."
    exit 1
}

$ExcelFilePath = $ExcelFiles[0].FullName

# Find the 'Documents' or 'Evidence' folder
$DocFolder = Join-Path $CurrentDir 'Documents'
if (-not (Test-Path $DocFolder)) {
    $DocFolder = Join-Path $CurrentDir 'Evidence'
    if (-not (Test-Path $DocFolder)) {
        Write-Error "Neither 'Documents' nor 'Evidence' folder was found in the current directory."
        exit 1
    }
}

# Get all files in the document folder
Write-Host "Enumerating files in '$($DocFolder)'..."
$Files = Get-ChildItem -Path $DocFolder -File

$TotalFiles = $Files.Count
$FileCounter = 0

# Create a hashtable for quick lookup of files by Document ID
$FileLookup = @{}
$DuplicateFileLookup = @{} # Create duplicate file listing
foreach ($File in $Files) {
    $FileCounter++
    if ($FileCounter % 5000 -eq 0) {
        Write-Host "$FileCounter of $TotalFiles files enumerated..."
    }

    # Extract the filename without extension
    $DocID = $File.BaseName.Trim().ToUpper()
    if (-not $FileLookup.ContainsKey($DocID)) {
        $FileLookup[$DocID] = $File.Name
        $DuplicateFileLookup[$DocID] = $File.Name # Populate duplicate structure
    } else {
        Write-Error "Multiple files found with the Document ID: $DocID"
        exit 1
    }
}

Write-Host "File enumeration completed. Total files found: $TotalFiles"

# Prepare hashtables for missing and extra files
$LinkedDocIDs = @{}
$MissingFiles = New-Object System.Collections.Generic.HashSet[string]
$TotalHyperlinksAdded = 0  # Initialize total hyperlinks counter

# Open Excel application
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $true
$Excel.DisplayAlerts = $false

# Open the Excel workbook
Write-Host "Opening Excel File"
$Workbook = $Excel.Workbooks.Open($ExcelFilePath)

try {
    # Get the count of worksheets
    $WorksheetCount = $Workbook.Worksheets.Count

    # Iterate through each worksheet using a for loop
    for ($w = 1; $w -le $WorksheetCount; $w++) {
        $Worksheet = $Workbook.Worksheets.Item($w)
        Write-Host "Processing worksheet: '$($Worksheet.Name)'"

        # Find the 'Document ID' column in the first three rows
        $DocIDColumn = $null
        $FoundHeader = $false

        for ($Row = 1; $Row -le 3; $Row++) {
            $UsedRange = $Worksheet.UsedRange
            $Rows = $UsedRange.Rows
            $RowRange = $Rows.Item($Row)
            $Columns = $RowRange.Columns

            $ColumnCount = $Columns.Count
            for ($Col = 1; $Col -le $ColumnCount; $Col++) {
                $Cell = $Columns.Item($Col)
                if ($Cell.Text -match '^\s*Document ID\s*$') {
                    $DocIDColumn = $Cell.Column
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
            Write-Host "No 'Document ID' column found in worksheet '$($Worksheet.Name)'. Skipping..."
            # Release $Worksheet
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Worksheet) | Out-Null
            $Worksheet = $null
            continue
        }

        Write-Host "Starting hyperlinking in worksheet '$($Worksheet.Name)'..."

        # Get the used range starting from the row after headers
        $UsedRange = $Worksheet.UsedRange
        $StartRow = $UsedRange.Row + 1
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
                    if ($LinkCounter % 5000 -eq 0) {
                        Write-Host "$LinkCounter hyperlinks created in worksheet '$($Worksheet.Name)'..."
                    }
                } else {
                    # Document ID found but no corresponding file
                    $MissingFiles.Add($DocID) | Out-Null
                }
            }
            # Release $Cell
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Cell) | Out-Null
            $Cell = $null
        }

        Write-Host "Finished hyperlinking in worksheet '$($Worksheet.Name)'. Total hyperlinks created: $LinkCounter"

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
}

# Create errata file with timestamp
$ErrataFilePath = Join-Path $CurrentDir ("errata_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".txt")

# Output extraneous files and missing Document IDs to errata file
Add-Content -Path $ErrataFilePath -Value "Extraneous Files:"
foreach ($ExtraFile in $DuplicateFileLookup.Keys) {
    Add-Content -Path $ErrataFilePath -Value "File: $($DuplicateFileLookup[$ExtraFile])"
}

Add-Content -Path $ErrataFilePath -Value "`nMissing Files:"
foreach ($MissingFile in $MissingFiles) {
    Add-Content -Path $ErrataFilePath -Value "Document ID: $MissingFile"
}

# Output confirmation message to console
Write-Host "Errata saved to $ErrataFilePath"
