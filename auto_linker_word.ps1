#
# V15 Prototype
# Stage 1: Evidence Source Discovery + Hashtable Build
#

# Folder names supported
$SupportedFolders = @("Evidence", "Documents")

Write-Host "Checking for evidence source folders..."

# Find matching folders in current directory
$EvidenceFolders = Get-ChildItem -Directory |
    Where-Object { $_.Name -in $SupportedFolders }

# Validation
if ($EvidenceFolders.Count -eq 0) {
    Write-Error "No 'Evidence' or 'Documents' folder found."
    exit 1
}

if ($EvidenceFolders.Count -gt 1) {
    Write-Error "Both 'Evidence' and 'Documents' folders were found. Only one may exist."
    exit 1
}

$EvidenceFolder = $EvidenceFolders[0]

Write-Host "Using evidence source folder:" -NoNewline
Write-Host " $($EvidenceFolder.Name)" -ForegroundColor Green

#
# Enumerate files
#

$EnumerationStart = Get-Date

$EvidenceFiles = Get-ChildItem -Path $EvidenceFolder.FullName -File

Write-Host "Found $($EvidenceFiles.Count) files."

#
# Build lookup structures
#

$EvidenceLookup = @{}
$UnreferencedEvidence = @{}

$FileCounter = 0
$CalloutInterval = 5000

$InvalidEvidenceFiles = @()

foreach ($File in $EvidenceFiles) {

    $FileCounter++

    if ($FileCounter % $CalloutInterval -eq 0) {
        Write-Host "    $FileCounter files indexed."
    }

    #
    # Filename without extension
    #

    $EvidenceID = $File.BaseName.Trim().ToUpper()

    #
    # Validate filename format
    #

    if ($EvidenceID -notmatch '^[A-Z]{3}\.\d{3}\.\d{3}\.\d{3,4}$') {

        $InvalidEvidenceFiles += $File.Name

        continue
    }

    #
    # Duplicate detection
    #

    if ($EvidenceLookup.ContainsKey($EvidenceID)) {

        Write-Host ""
        Write-Host "Duplicate evidence identifier detected:" -ForegroundColor Red
        Write-Host "    $EvidenceID"
        Write-Host ""
        Write-Host "Existing:"
        Write-Host "    $($EvidenceLookup[$EvidenceID])"
        Write-Host ""
        Write-Host "Duplicate:"
        Write-Host "    $($File.Name)"
        Write-Host ""

        return
    }

    $EvidenceLookup[$EvidenceID] = $File.Name

    #
    # Clone tracking structure for later reconciliation
    #

    $UnreferencedEvidence[$EvidenceID] = $File.Name
}

#
# Report invalid filenames
#

if ($InvalidEvidenceFiles.Count -gt 0) {

    Write-Host ""
    Write-Host "Invalid evidence filenames found:" -ForegroundColor Red
    Write-Host "    $($InvalidEvidenceFiles.Count)"
    Write-Host ""

    foreach ($InvalidFile in ($InvalidEvidenceFiles | Sort-Object)) {
        Write-Host "        $InvalidFile"
    }

    Write-Host ""
    Write-Host "Expected filename format:" -ForegroundColor Yellow
    Write-Host "    ABC.123.123.123.ext"
    Write-Host "    ABC.123.123.1234.ext"
    Write-Host ""
    Write-Host "No additional text is permitted in the filename."
    Write-Host ""

    return
}

$EnumerationDuration = (Get-Date) - $EnumerationStart

Write-Host ""
Write-Host "Evidence indexing complete."
Write-Host "    Indexed IDs: $($EvidenceLookup.Count)"
Write-Host "    Duration: $($EnumerationDuration.ToString('hh\:mm\:ss'))"

#
# Enumerate DOCX files in current directory
#

$WordDocuments = Get-ChildItem -File -Filter "*.docx" |
    Where-Object { $_.Name -notlike "~$*" } |
    Sort-Object Name

if ($WordDocuments.Count -eq 0) {
    Write-Error "No .docx files found in the current directory."
    exit 1
}

#
# Auto-select if only one document exists
#

if ($WordDocuments.Count -eq 1)
{
    $SelectedDocument = $WordDocuments[0]

    Write-Host ""
    Write-Host "Single Word document found. Auto-selecting:" -NoNewline
    Write-Host " $($SelectedDocument.Name)" -ForegroundColor Green
}
else
{
    Write-Host ""
    Write-Host "Available Word documents:"
    Write-Host ""

    for ($i = 0; $i -lt $WordDocuments.Count; $i++)
    {
        Write-Host "[$($i + 1)] $($WordDocuments[$i].Name)"
    }

    Write-Host ""

    do
    {
        $Selection = Read-Host "Enter document number"

        $ValidSelection = (
            $Selection -match '^\d+$' -and
            [int]$Selection -ge 1 -and
            [int]$Selection -le $WordDocuments.Count
        )

        if (-not $ValidSelection)
        {
            Write-Host "Invalid selection. Please try again." -ForegroundColor Yellow
        }

    } until ($ValidSelection)

    $SelectedDocument = $WordDocuments[[int]$Selection - 1]
}

Write-Host ""
Write-Host "Selected document:" -NoNewline
Write-Host " $($SelectedDocument.Name)" -ForegroundColor Green

$WordDocumentPath = $SelectedDocument.FullName

#
# Open Word document
#

Write-Host ""
Write-Host "Opening Word document..."

$Word = $null
$Document = $null

try
{
    $Word = New-Object -ComObject Word.Application
    $Word.Visible = $false
    $Word.DisplayAlerts = 0

    $Document = $Word.Documents.Open($WordDocumentPath)

    Write-Host "Document opened successfully." -ForegroundColor Green

#
# Extract document text into memory
#

Write-Host ""
Write-Host "Extracting document text..."

#
# Main document body
#

$BodyText = $Document.Content.Text

#
# Footnotes
#

$FootnoteTextBuilder = New-Object System.Text.StringBuilder

foreach ($Footnote in $Document.Footnotes)
{
    if ($Footnote.Range.Text)
    {
        $null = $FootnoteTextBuilder.AppendLine($Footnote.Range.Text)
    }
}

$FootnoteText = $FootnoteTextBuilder.ToString()

#
# Statistics
#

$FootnoteCount = $Document.Footnotes.Count

$BodyCharCount = $BodyText.Length
$FootnoteCharCount = $FootnoteText.Length
$TotalCharCount = $BodyCharCount + $FootnoteCharCount

Write-Host ""
Write-Host "Document text extraction complete."

Write-Host "    Footnotes: " -NoNewline
Write-Host $FootnoteCount -ForegroundColor Green

Write-Host "    Main document characters: " -NoNewline
Write-Host $BodyCharCount -ForegroundColor Green

Write-Host "    Footnote characters: " -NoNewline
Write-Host $FootnoteCharCount -ForegroundColor Green

Write-Host "    Total characters: " -NoNewline
Write-Host $TotalCharCount -ForegroundColor Green

#
# Main document Bates scan
#

$ScanStart = Get-Date

$BatesPattern = '(?<!\S)[A-Z]{3}\.\d{3}\.\d{3}\.\d{3,4}(?!\S)'

$BodyMatches = [System.Text.RegularExpressions.Regex]::Matches(
    $BodyText,
    $BatesPattern
)

$BodyBatesLookup = @{}

foreach ($Match in $BodyMatches)
{
    $BodyBatesLookup[$Match.Value.ToUpper()] = $true
}

$ScanDuration = (Get-Date) - $ScanStart

Write-Host ""
Write-Host "Main document Bates scan complete."

Write-Host "    Occurrences: " -NoNewline
Write-Host $BodyMatches.Count -ForegroundColor Green

Write-Host "    Unique Bates references: " -NoNewline
Write-Host $BodyBatesLookup.Count -ForegroundColor Green

Write-Host "    Duration: $($ScanDuration.ToString('hh\:mm\:ss'))"

#
# Footnote Bates scan
#

$ScanStart = Get-Date

$FootnoteMatches = [System.Text.RegularExpressions.Regex]::Matches(
    $FootnoteText,
    $BatesPattern
)

$FootnoteBatesLookup = @{}

foreach ($Match in $FootnoteMatches)
{
    $FootnoteBatesLookup[$Match.Value.ToUpper()] = $true
}

$ScanDuration = (Get-Date) - $ScanStart

Write-Host ""
Write-Host "Footnote Bates scan complete."

Write-Host "    Occurrences: " -NoNewline
Write-Host $FootnoteMatches.Count -ForegroundColor Green

Write-Host "    Unique Bates references: " -NoNewline
Write-Host $FootnoteBatesLookup.Count -ForegroundColor Green

Write-Host "    Duration: $($ScanDuration.ToString('hh\:mm\:ss'))"

#
# Combined Bates references
#

$AllBatesLookup = @{}

foreach ($Bates in $BodyBatesLookup.Keys)
{
    $AllBatesLookup[$Bates] = $true
}

foreach ($Bates in $FootnoteBatesLookup.Keys)
{
    $AllBatesLookup[$Bates] = $true
}

Write-Host ""
Write-Host "Document Bates summary."

Write-Host "    Total unique Bates references: " -NoNewline
Write-Host $AllBatesLookup.Count -ForegroundColor Green

#
# Reconciliation
#

$MatchedCount = 0

$MissingFromFolder = @{}
$UnreferencedEvidence = @{}

#
# Copy all evidence into the
# unreferenced bucket initially
#

foreach ($EvidenceID in $EvidenceLookup.Keys)
{
    $UnreferencedEvidence[$EvidenceID] = $true
}

#
# Compare document references
# against evidence folder
#

foreach ($Bates in $AllBatesLookup.Keys)
{
    if ($EvidenceLookup.ContainsKey($Bates))
    {
        $MatchedCount++

        #
        # Remove matched evidence
        # from unreferenced list
        #

        $UnreferencedEvidence.Remove($Bates)
    }
    else
    {
        #
        # In document but not folder
        #

        $MissingFromFolder[$Bates] = $true
    }
}

Write-Host ""
Write-Host "Evidence reconciliation complete."

Write-Host "    Matched references: " -NoNewline
Write-Host $MatchedCount -ForegroundColor Green

Write-Host "    Referenced but missing evidence (in document but no matching file in folder): " -NoNewline
Write-Host $MissingFromFolder.Count -ForegroundColor Yellow

if ($MissingFromFolder.Count -gt 0)
{
    foreach ($Bates in ($MissingFromFolder.Keys | Sort-Object))
    {
        Write-Host "        $Bates"
    }
}

Write-Host "    Evidence not referenced (file in folder but not referred to in document): " -NoNewline
Write-Host $UnreferencedEvidence.Count -ForegroundColor Yellow

if ($UnreferencedEvidence.Count -gt 0)
{
    foreach ($Bates in ($UnreferencedEvidence.Keys | Sort-Object))
    {
        Write-Host "        $Bates"
    }
}
##############################

#
# Hyperlink all matched Bates references
#

$MatchedBates = $AllBatesLookup.Keys |
    Where-Object { $EvidenceLookup.ContainsKey($_) } |
    Sort-Object

if ($MatchedBates.Count -gt 0)
{
	$HyperlinkStart = Get-Date
    Write-Host ""
    Write-Host "Starting hyperlinking..."

    $MatchedBatesProcessed = 0
    $BodyLinksAdded = 0
    $FootnoteLinksAdded = 0

	foreach ($CurrentBates in $MatchedBates)
	{
		$CurrentBatesLinks = 0

		Write-Host ("Searching for {0}" -f $CurrentBates)
		$BatesStart = Get-Date
        $TargetPath = "./" + $EvidenceFolder.Name + "/" + $EvidenceLookup[$CurrentBates]
        Write-Host "        $TargetPath"

        #
        # BODY
        #

        $BodyRange = $Document.Content.Duplicate

        $Find = $BodyRange.Find

        $Find.ClearFormatting()
        $Find.Text = $CurrentBates
        $Find.Forward = $true
        $Find.Wrap = 0

		$SafetyCounter = 0

		while ($Find.Execute())
		{
			Write-Host (" {0}-{1}" -f $FoundRange.Start, $FoundRange.End)
			$FoundRange = $SearchRange.Duplicate
			$Document.Hyperlinks.Add(
				$BodyRange,
				$TargetPath
			) | Out-Null
			
			$SearchRange.Start = $FoundRange.End
			$SearchRange.End = $SearchScope.End

			$BodyLinksAdded++

			$SafetyCounter++

			if ($SafetyCounter -gt 20)
			{
				Write-Host "        Safety break triggered in body search for $CurrentBates" -ForegroundColor Yellow
				break
			}

			#
			# Move past the current match
			#

			$BodyRange.Collapse(1)
		}

        #
        # FOOTNOTES
        #
		
		$SafetyCounter = 0
		Write-Host (" {0}-{1}" -f $FoundRange.Start, $FoundRange.End)
		$FoundRange = $SearchRange.Duplicate
		while ($Find.Execute())
		{
			$Document.Hyperlinks.Add(
				$Range,
				$TargetPath
			) | Out-Null
			$SearchRange.Start = $FoundRange.End
			$SearchRange.End = $SearchScope.End
			$FootnoteLinksAdded++
			$CurrentBatesLinks++

			$SafetyCounter++

			if ($SafetyCounter -gt 20)
			{
				Write-Host "        Safety break triggered in footnote search for $CurrentBates" -ForegroundColor Yellow
				break
			}

			#
			# Move past the current match
			#

			$Range.Collapse(1)
		}
		Write-Host " $CurrentBates : $CurrentBatesLinks"

        $MatchedBatesProcessed++
		
		$BatesDuration = (Get-Date) - $BatesStart
		Write-Host "        $CurrentBates : $CurrentLinksAdded links in $($BatesDuration.TotalSeconds.ToString('0.00')) seconds"

        Write-Host "    $MatchedBatesProcessed / $($MatchedBates.Count) Bates linked."
		
		# TESTING LIMIT ITERATION OF HYPERLINKING
        #break
    }

	$HyperlinkDuration = (Get-Date) - $HyperlinkStart
    Write-Host ""
    Write-Host "Hyperlinking complete."
	Write-Host " Hyperlinking duration: " -NoNewline

Write-Host $HyperlinkDuration.ToString('hh\:mm\:ss') -ForegroundColor Green

    Write-Host "    Matched Bates processed: " -NoNewline
    Write-Host $MatchedBatesProcessed -ForegroundColor Green

    Write-Host "    Body hyperlinks added: " -NoNewline
    Write-Host $BodyLinksAdded -ForegroundColor Green

    Write-Host "    Footnote hyperlinks added: " -NoNewline
    Write-Host $FootnoteLinksAdded -ForegroundColor Green

    Write-Host "    Total hyperlinks added: " -NoNewline
    Write-Host ($BodyLinksAdded + $FootnoteLinksAdded) -ForegroundColor Green
}
else
{
    Write-Host ""
    Write-Host "No matched Bates references available for hyperlinking." -ForegroundColor Yellow
}


###############################
##
## Catch Errors and Close Document
##
}
catch
{
    Write-Error $_
}
finally
{
    Write-Host ""

	if ($Document)
	{
		$TotalLinksAdded = $BodyLinksAdded + $FootnoteLinksAdded

		if ($TotalLinksAdded -gt 1)
		{
			Write-Host "Saving document..."

			$Document.Save()

			Write-Host "Document saved." -ForegroundColor Green
		}
		else
		{
			Write-Host "Document not saved because fewer than two hyperlinks were created." -ForegroundColor Yellow
		}
	}

    Write-Host ""

    Write-Host "Closing Word document..."

    if ($Document)
    {
        $Document.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Document) | Out-Null
        $Document = $null
    }

    if ($Word)
    {
        $Word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Word) | Out-Null
        $Word = $null
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Host "Word closed." -ForegroundColor Green
}