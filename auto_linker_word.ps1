#
# V15 Prototype
# Stage 1: Evidence Source Discovery + Hashtable Build
#

#
# Pure functions (no script state, no I/O) — testable without Word installed
#

function Get-BatesMatches
{
    param(
        [string]$Text,
        [string]$Pattern
    )

    return [System.Text.RegularExpressions.Regex]::Matches(
        $Text,
        $Pattern,
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )
}

function Get-BatesLookupFromMatches
{
    param(
        $Matches
    )

    $Lookup = @{}

    foreach ($Match in $Matches)
    {
        $Lookup[$Match.Value.ToUpper()] = $true
    }

    return $Lookup
}

function Merge-BatesLookup
{
    param(
        [hashtable[]]$Lookups
    )

    $Merged = @{}

    foreach ($Lookup in $Lookups)
    {
        foreach ($Bates in $Lookup.Keys)
        {
            $Merged[$Bates] = $true
        }
    }

    return $Merged
}

function Get-BatesReconciliation
{
    param(
        [hashtable]$DocumentBatesLookup,
        [hashtable]$EvidenceLookup
    )

    $MatchedCount = 0
    $MissingFromFolder = @{}
    $UnreferencedEvidence = @{}

    foreach ($EvidenceID in $EvidenceLookup.Keys)
    {
        $UnreferencedEvidence[$EvidenceID] = $true
    }

    foreach ($Bates in $DocumentBatesLookup.Keys)
    {
        if ($EvidenceLookup.ContainsKey($Bates))
        {
            $MatchedCount++
            $UnreferencedEvidence.Remove($Bates)
        }
        else
        {
            $MissingFromFolder[$Bates] = $true
        }
    }

    return [PSCustomObject]@{
        MatchedCount         = $MatchedCount
        MissingFromFolder    = $MissingFromFolder
        UnreferencedEvidence = $UnreferencedEvidence
    }
}

#
# Word COM constants (avoids unnamed magic numbers at call sites)
#

$wdFootnotesStory = 2
$wdCollapseEnd = 0
$wdFindStop = 0

#
# Finds every occurrence of $SearchText within $SearchRange and hyperlinks
# each one to $TargetPath. Shared by the body and footnote passes below so
# the Find/Collapse/reverse-apply logic only exists once.
#

function Add-BatesHyperlinksToRange
{
    param(
        $Document,
        $SearchRange,
        [string]$SearchText,
        [string]$TargetPath,
        [int]$SafetyLimit = 20
    )

    $Find = $SearchRange.Find

    $Find.ClearFormatting()
    $Find.Text = $SearchText
    $Find.Forward = $true
    $Find.Wrap = $wdFindStop

    $MatchRanges = @()
    $SafetyCounter = 0

    while ($Find.Execute())
    {
        $MatchRanges += , @($SearchRange.Start, $SearchRange.End)

        $SafetyCounter++

        if ($SafetyCounter -gt $SafetyLimit)
        {
            Write-Host "        Safety break triggered while searching for $SearchText" -ForegroundColor Yellow
            break
        }

        #
        # Move past the current match before searching again
        #

        $SearchRange.Collapse($wdCollapseEnd)
    }

    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Find) | Out-Null

    #
    # Matches are located first, then hyperlinks are applied last-to-first.
    # Applying in place while still searching would let Find re-match the Bates
    # ID hidden inside the newly inserted hyperlink field's address text.
    #

    for ($MatchIndex = $MatchRanges.Count - 1; $MatchIndex -ge 0; $MatchIndex--)
    {
        $MatchStart, $MatchEnd = $MatchRanges[$MatchIndex]

        $HyperlinkRange = $SearchRange.Duplicate
        $HyperlinkRange.Start = $MatchStart
        $HyperlinkRange.End = $MatchEnd

        $Document.Hyperlinks.Add(
            $HyperlinkRange,
            $TargetPath
        ) | Out-Null

        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($HyperlinkRange) | Out-Null
    }

    return $MatchRanges.Count
}

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
    # Optional "-N"/"-NN" suffix supported (e.g. ABX.456.876.0000-1)
    #

    if ($EvidenceID -notmatch '^[A-Z]{3}\.\d{3}\.\d{3}\.\d{3,4}(-\d{1,2})?$') {

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

        exit 1
    }

    $EvidenceLookup[$EvidenceID] = $File.Name
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
    Write-Host "    ABC.123.123.1234-1.ext"
    Write-Host "    ABC.123.123.1234-12.ext"
    Write-Host ""
    Write-Host "No additional text is permitted in the filename."
    Write-Host ""

    exit 1
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
# Scope is intentionally limited to the main body and footnotes. Headers,
# footers, endnotes, comments, and text boxes are not scanned or hyperlinked.
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

# Boundaries reject adjacent letters/digits (avoids partial matches) but allow punctuation like ",./)"
# Optional "-N"/"-NN" suffix supported (e.g. ABX.456.876.0000-1)
$BatesPattern = '(?<![A-Za-z0-9])[A-Z]{3}\.\d{3}\.\d{3}\.\d{3,4}(-\d{1,2})?(?![A-Za-z0-9])'

$BodyMatches = Get-BatesMatches -Text $BodyText -Pattern $BatesPattern
$BodyBatesLookup = Get-BatesLookupFromMatches -Matches $BodyMatches

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

$FootnoteMatches = Get-BatesMatches -Text $FootnoteText -Pattern $BatesPattern
$FootnoteBatesLookup = Get-BatesLookupFromMatches -Matches $FootnoteMatches

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

$AllBatesLookup = Merge-BatesLookup -Lookups @($BodyBatesLookup, $FootnoteBatesLookup)

Write-Host ""
Write-Host "Document Bates summary."

Write-Host "    Total unique Bates references: " -NoNewline
Write-Host $AllBatesLookup.Count -ForegroundColor Green

#
# Reconciliation
#

$Reconciliation = Get-BatesReconciliation -DocumentBatesLookup $AllBatesLookup -EvidenceLookup $EvidenceLookup

$MatchedCount = $Reconciliation.MatchedCount
$MissingFromFolder = $Reconciliation.MissingFromFolder
$UnreferencedEvidence = $Reconciliation.UnreferencedEvidence

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
# Only the main body and footnote stories are searched here, matching the
# scan scope above. Word's Find engine is used instead of .NET regex because
# Hyperlinks.Add requires a live Range, not a plain string offset.
#

$MatchedBates = $AllBatesLookup.Keys |
    Where-Object { $EvidenceLookup.ContainsKey($_) } |
    Sort-Object

$MatchedBatesProcessed = 0
$BodyLinksAdded = 0
$FootnoteLinksAdded = 0

if ($MatchedBates.Count -gt 0)
{
	$HyperlinkStart = Get-Date
    Write-Host ""
    Write-Host "Starting hyperlinking..."

	foreach ($CurrentBates in $MatchedBates)
	{
		$BatesStart = Get-Date
        $TargetPath = "./" + $EvidenceFolder.Name + "/" + $EvidenceLookup[$CurrentBates]

        #
        # BODY
        #

        $BodyRange = $Document.Content.Duplicate

        $BodyMatchCount = Add-BatesHyperlinksToRange -Document $Document -SearchRange $BodyRange -SearchText $CurrentBates -TargetPath $TargetPath

        $BodyLinksAdded += $BodyMatchCount

        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($BodyRange) | Out-Null

        #
        # FOOTNOTES
        #

        $FootnoteMatchCount = 0

        if ($FootnoteCount -gt 0)
        {
            $FootnoteRange = $Document.StoryRanges.Item($wdFootnotesStory)

            if ($FootnoteRange)
            {
                $FootnoteMatchCount = Add-BatesHyperlinksToRange -Document $Document -SearchRange $FootnoteRange -SearchText $CurrentBates -TargetPath $TargetPath

                $FootnoteLinksAdded += $FootnoteMatchCount

                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($FootnoteRange) | Out-Null
            }
        }

        $CurrentBatesLinks = $BodyMatchCount + $FootnoteMatchCount
        $BatesDuration = (Get-Date) - $BatesStart
        $MatchedBatesProcessed++

        Write-Host ("    [{0,2}/{1}] {2,-20}" -f $MatchedBatesProcessed, $MatchedBates.Count, $CurrentBates) -NoNewline
        Write-Host (" {0,2} link{1}" -f $CurrentBatesLinks, $(if ($CurrentBatesLinks -eq 1) { "" } else { "s" })) -NoNewline -ForegroundColor Green
        Write-Host (" ({0} body, {1} footnote, {2}s)" -f $BodyMatchCount, $FootnoteMatchCount, $BatesDuration.TotalSeconds.ToString('0.00'))
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

		# Save only if at least one hyperlink was actually created
		if ($TotalLinksAdded -ge 1)
		{
			Write-Host "Saving document..."

			# Frees the undo history built up by the hyperlink insertions before saving
			$Document.UndoClear()

			$Document.Save()

			Write-Host "Document saved." -ForegroundColor Green
		}
		else
		{
			Write-Host "Document not saved because no hyperlinks were created." -ForegroundColor Yellow
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