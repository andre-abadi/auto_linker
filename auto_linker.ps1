# Get the current directory where the script is being executed
$currentDirectory = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Record the start time for errata file naming
$scriptStartTime = Get-Date
$datetimeString = $scriptStartTime.ToString("yyyyMMdd_HHmmss")

# Initialize lists for missing documents and extraneous files
$missingDocIDs = @()
$extraneousFiles = @()

# Find all Excel files in the current directory
$excelFiles = Get-ChildItem -Path $currentDirectory -Filter *.xlsx -File

# Error handling for Excel files
if ($excelFiles.Count -eq 0) {
    throw "No Excel file found in the current directory."
}
elseif ($excelFiles.Count -gt 1) {
    throw "Multiple Excel files found in the current directory. Please ensure only one Excel file is present."
}
else {
    $excelFile = $excelFiles[0]
    $excelFilePath = $excelFile.FullName
    Write-Host "Excel file detected: '$($excelFile.Name)'"
}

# Get the parent directory of the Excel file (same as current directory)
$parentFolder = $currentDirectory

# Paths to potential folders
$documentsPath = Join-Path -Path $parentFolder -ChildPath "Documents"
$evidencePath = Join-Path -Path $parentFolder -ChildPath "Evidence"

# Check if the folders exist
$documentsExists = Test-Path -Path $documentsPath
$evidenceExists = Test-Path -Path $evidencePath

# Determine which folder to use
if ($documentsExists -and $evidenceExists) {
    throw "Both 'Documents' and 'Evidence' folders exist. Please ensure only one is present."
}
elseif ($documentsExists) {
    $targetFolder = $documentsPath
    Write-Host "Using folder: 'Documents'"
}
elseif ($evidenceExists) {
    $targetFolder = $evidencePath
    Write-Host "Using folder: 'Evidence'"
}
else {
    throw "Neither 'Documents' nor 'Evidence' folder found in the current directory."
}

# Gather all files in the target folder into a hashtable
Write-Host "Gathering files from the folder into memory..."
$filesHashtable = @{}
$allFilenames = @()
$fileCounter = 0
$updateInterval = 1000  # Update progress every 1000 files

Get-ChildItem -Path $targetFolder -File | ForEach-Object {
    $filenameWithoutExtension = $_.BaseName
    $allFilenames += $filenameWithoutExtension
    if (-not $filesHashtable.ContainsKey($filenameWithoutExtension)) {
        $filesHashtable[$filenameWithoutExtension] = $_.FullName
    }
    else {
        # Ignore duplicates for performance
    }

    $fileCounter++

    if (($fileCounter % $updateInterval) -eq 0) {
        Write-Host "Files read into memory: $fileCounter"
    }
}

# Output total files collected
Write-Host "Total files collected: $($filesHashtable.Count)"

# Create an Excel COM object
$excel = New-Object -ComObject Excel.Application

# Performance settings
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false
$excel.EnableEvents = $false
$excel.AskToUpdateLinks = $false
$excel.EnableAutoComplete = $false

try {
    # Open the workbook
    $workbook = $excel.Workbooks.Open($excelFilePath)

    # Additional performance settings for data connections
    foreach ($connection in $workbook.Connections) {
        $connection.OLEDBConnection.BackgroundQuery = $false
    }

    # Initialize hyperlink counter
    $totalHyperlinksAdded = 0

    # Get the total number of worksheets
    $totalWorksheets = $workbook.Worksheets.Count
    $worksheetIndex = 0

    # HashSet to store DocIDs found in Excel
    $docIDsInExcel = New-Object System.Collections.Generic.HashSet[string]

    # Loop through each worksheet in the workbook
    foreach ($worksheet in $workbook.Worksheets) {
        $worksheetIndex++
        Write-Host "Processing worksheet ($worksheetIndex/$totalWorksheets): $($worksheet.Name)"

        $sheetHyperlinksAdded = 0  # Reset counter for each new sheet

        # Check for and refresh query tables
        if ($worksheet.QueryTables.Count -gt 0) {
            Write-Host "Worksheet '$($worksheet.Name)' contains query tables. Refreshing data..."
            foreach ($queryTable in $worksheet.QueryTables) {
                try {
                    $queryTable.Refresh($false)
                }
                catch {
                    Write-Warning "Could not refresh query table in worksheet '$($worksheet.Name)': $_"
                }
            }
        }

        # Process tables first
        $tables = $worksheet.ListObjects
        $docIdColumnIndex = 0  # Reset for each table
        
        if ($tables.Count -gt 0) {
            # Process each table in the worksheet
            foreach ($table in $tables) {
                $headerRow = $table.HeaderRowRange
                $dataRange = $table.DataBodyRange
                
                # Find Document ID column in table
                $headerValues = $headerRow.Value2
                for ($col = 1; $col -le $headerValues.GetLength(1); $col++) {
                    $cellValue = $headerValues[1, $col]
                    if ($cellValue -eq "Document ID") {
                        $docIdColumnIndex = $col
                        Write-Host "Found 'Document ID' column at position $docIdColumnIndex in table '$($table.Name)'"
                        break
                    }
                }

                if ($docIdColumnIndex -gt 0) {
                    # Get DocIDs from the table column
                    $docIds = $dataRange.Columns($docIdColumnIndex).Value2
                    
                    if ($docIds -eq $null) {
                        Write-Host "No DocIDs found in table '$($table.Name)'. Skipping."
                        continue
                    }

                    # Prepare arrays for hyperlink addresses and display texts
                    $hyperlinkAddresses = @()
                    $hyperlinkDisplayTexts = @()
                    # Collect cells that need hyperlinks
                    $cellsToHyperlink = @()

                    # Get the number of rows in the DocIDs array
                    if ($docIds.GetType().IsArray) {
                        $rowCount = $docIds.GetLength(0)
                    }
                    else {
                        # Only one DocID present
                        $rowCount = 1
                        $docIds = @(@($docIds))
                    }

                    for ($i = 1; $i -le $rowCount; $i++) {
                        $docId = $docIds[$i, 1]
                        if ($docId -ne $null -and $docId -ne "") {
                            # Add DocID to the set of DocIDs found in Excel
                            $docIDsInExcel.Add($docId) | Out-Null

                            if ($filesHashtable.ContainsKey($docId)) {
                                $filePath = $filesHashtable[$docId]
                                # Store the cell and hyperlink info
                                $cell = $dataRange.Cells($i, $docIdColumnIndex)
                                $cellsToHyperlink += $cell
                                $hyperlinkAddresses += $filePath
                                $hyperlinkDisplayTexts += $docId
                            }
                            else {
                                # Collect missing DocIDs
                                $missingDocIDs += $docId
                            }

                            $totalHyperlinksAdded++
                            $sheetHyperlinksAdded++
                            if (($totalHyperlinksAdded % $updateInterval) -eq 0) {
                                Write-Host "Total hyperlinks processed: $totalHyperlinksAdded"
                            }
                        }
                    }
                }
            }
            # Skip regular range processing if we processed any tables with Document ID columns
            if ($docIdColumnIndex -gt 0) {
                Write-Host "Completed processing table(s) in worksheet '$($worksheet.Name)' - Added $sheetHyperlinksAdded hyperlinks in this sheet"
                continue
            }
        }

        # Then process regular ranges
        $usedRange = $worksheet.UsedRange

        if ($usedRange -eq $null) {
            Write-Host "Worksheet '$($worksheet.Name)' is empty. Skipping."
            continue
        }

        # Find the header row index (assuming it's the first row)
        $headerRowIndex = 1  # Adjust this if your headers are in a different row
        $headerRow = $worksheet.Rows.Item($headerRowIndex)
        $docIdColumnIndex = 0

        # Find the column index for "Document ID"
        $headerValues = $headerRow.Value2

        if ($headerValues -eq $null) {
            Write-Host "Header row is empty in worksheet '$($worksheet.Name)'. Skipping."
            continue
        }

        # Get the number of columns in the header row
        $colCount = $headerValues.GetLength(1)

        for ($col = 1; $col -le $colCount; $col++) {
            $cellValue = $headerValues[1, $col]
            if ($cellValue -eq "Document ID") {
                $docIdColumnIndex = $col
                Write-Host "Found 'Document ID' column at position $docIdColumnIndex in worksheet '$($worksheet.Name)'"
                break
            }
        }

        if ($docIdColumnIndex -eq 0) {
            Write-Host "Document ID column not found in worksheet '$($worksheet.Name)'. Skipping."
            continue
        }

        # Get all DocIDs from the column into an array
        $lastRow = $usedRange.Rows.Count
        $docIdRange = $worksheet.Range(
            $worksheet.Cells.Item($headerRowIndex + 1, $docIdColumnIndex),
            $worksheet.Cells.Item($lastRow, $docIdColumnIndex)
        )
        $docIds = $docIdRange.Value2

        if ($docIds -eq $null) {
            Write-Host "No DocIDs found in worksheet '$($worksheet.Name)'. Skipping."
            continue
        }

        # Prepare arrays for hyperlink addresses and display texts
        $hyperlinkAddresses = @()
        $hyperlinkDisplayTexts = @()
        # Collect cells that need hyperlinks
        $cellsToHyperlink = @()

        # Get the number of rows in the DocIDs array
        if ($docIds.GetType().IsArray) {
            $rowCount = $docIds.GetLength(0)
        }
        else {
            # Only one DocID present
            $rowCount = 1
            $docIds = @(@($docIds))
        }

        for ($i = 1; $i -le $rowCount; $i++) {
            $docId = $docIds[$i, 1]
            if ($docId -ne $null -and $docId -ne "") {
                # Add DocID to the set of DocIDs found in Excel
                $docIDsInExcel.Add($docId) | Out-Null

                if ($filesHashtable.ContainsKey($docId)) {
                    $filePath = $filesHashtable[$docId]
                    # Store the cell and hyperlink info
                    $cellRowIndex = $headerRowIndex + $i
                    $cell = $worksheet.Cells.Item($cellRowIndex, $docIdColumnIndex)
                    $cellsToHyperlink += $cell
                    $hyperlinkAddresses += $filePath
                    $hyperlinkDisplayTexts += $docId
                }
                else {
                    # Collect missing DocIDs
                    $missingDocIDs += $docId
                }

                $totalHyperlinksAdded++
                Write-Host "Processing hyperlink $totalHyperlinksAdded : $docId"
            }
        }

        # Add hyperlinks in bulk
        for ($j = 0; $j -lt $cellsToHyperlink.Count; $j++) {
            $cell = $cellsToHyperlink[$j]
            $address = $hyperlinkAddresses[$j]
            $displayText = $hyperlinkDisplayTexts[$j]

            # Add the hyperlink
            $cell.Hyperlinks.Add($cell, $address, $null, $null, $displayText)
        }
    }

    # Final progress update
    Write-Host "Total hyperlinks processed: $totalHyperlinksAdded"

    # Collect extraneous files (files not referenced in Excel)
    $extraneousFiles = $allFilenames | Where-Object { -not $docIDsInExcel.Contains($_) }

    # Summary of missing and surplus files
    Write-Host "Missing files count: $($missingDocIDs.Count)"
    Write-Host "Surplus files count: $($extraneousFiles.Count)"

    # If there are missing documents or extraneous files, create errata file
    if ($missingDocIDs.Count -gt 0 -or $extraneousFiles.Count -gt 0) {
        $errataFilename = "errata_$datetimeString.txt"
        $errataFilePath = Join-Path -Path $currentDirectory -ChildPath $errataFilename

        $errataContent = @()
        $errataContent += "Errata Report - Generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        $errataContent += ""

        if ($missingDocIDs.Count -gt 0) {
            $errataContent += "Missing Documents (DocIDs in Excel but no corresponding file):"
            $errataContent += $missingDocIDs | Sort-Object
            $errataContent += ""
        }

        if ($extraneousFiles.Count -gt 0) {
            $errataContent += "Extraneous Files (Files in folder but no corresponding DocID in Excel):"
            $errataContent += $extraneousFiles | Sort-Object
            $errataContent += ""
        }

        $errataContent | Set-Content -Path $errataFilePath -Encoding UTF8

        Write-Host "Errata file created: $errataFilename"
    }

    # Save and close the workbook
    $workbook.Save()
    $workbook.Close()
}
catch {
    Write-Error $_.Exception.Message
}
finally {
    # Restore Excel settings
    $excel.ScreenUpdating = $true
    $excel.EnableEvents = $true

    # Quit Excel application
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

# Clean up COM objects
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host "Hyperlinks added successfully."