# --- Script metadata ---
# Stage 1: Evidence source discovery and hashtable build.

# Folder containing the Word documents.
$TargetDir = Join-Path (Get-Location).ProviderPath "Covering Material" 

# Supported source folder names.
$SupportedFolders = @("Evidence", "Documents")

# Errata Outputs
$NoBatesPath = Join-Path (Get-Location).ProviderPath "_no_bates.txt"
$NoReferencePath = Join-Path (Get-Location).ProviderPath "_no_reference.txt"
$NoFilesPath = Join-Path (Get-Location).ProviderPath "_no_files.txt"


# Logging folder (optional)
$LogFolder = ""

# Start timer
$ScriptStart = Get-Date

function Write-ListingFile
{
    param(
        [string]$Path,
        [object[]]$Items
    )

    if ($null -eq $Items)
    {
        $Items = @()
    }

    Remove-Item -Path $Path -Force -ErrorAction SilentlyContinue
    [System.IO.File]::WriteAllLines($Path, [string[]]@($Items))
}

function Write-ExecutionLog
{
    param(
        [string]$FolderPath,
        [int]$HyperlinksCreated
    )

    if ([string]::IsNullOrWhiteSpace($FolderPath))
    {
        return
    }

    try
    {
        $ExistingFolder = Get-Item -LiteralPath $FolderPath -ErrorAction SilentlyContinue

        if ($ExistingFolder)
        {
            if (-not $ExistingFolder.PSIsContainer)
            {
                throw "The configured log path is not a folder: $FolderPath"
            }
        }
        else
        {
            New-Item -Path $FolderPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }

        $LogPath = Join-Path $FolderPath "auto_linker_word_log.csv"
        $LogEntry = [PSCustomObject]@{
            ExecutionTime     = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            Username          = $env:USERNAME
            HyperlinksCreated = $HyperlinksCreated
        }

        $ExistingLog = Get-Item -LiteralPath $LogPath -ErrorAction SilentlyContinue

        if ($ExistingLog)
        {
            if ($ExistingLog.PSIsContainer)
            {
                throw "The execution log path is a folder, not a CSV file: $LogPath"
            }

            $ExistingHeader = Get-Content -LiteralPath $LogPath -TotalCount 1 -ErrorAction Stop
            $NormalizedHeader = $ExistingHeader.Trim().TrimStart([char]0xFEFF) -replace '"', ''

            if (-not [string]::IsNullOrWhiteSpace($NormalizedHeader) -and
                $NormalizedHeader -ne "ExecutionTime,Username,HyperlinksCreated")
            {
                throw "The existing execution log has unexpected CSV columns: $LogPath"
            }

            if ([string]::IsNullOrWhiteSpace($NormalizedHeader))
            {
                $LogEntry | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
            }
            else
            {
                $LogEntry | Export-Csv -Path $LogPath -Append -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
            }
        }
        else
        {
            $LogEntry | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        }

    }
    catch
    {
        return
    }
}



# --- Pure functions ---
# These functions have no script state or I/O and can be tested without Word installed.

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

# --- Word COM constants ---
# These constants avoid unnamed magic numbers at call sites.

$wdFootnotesStory = 2
$wdCollapseEnd = 0
$wdFindStop = 0

# --- Hyperlink range helper ---
# Finds every occurrence of $SearchText in $SearchRange and links it to $TargetPath.
# The body and footnote passes share this Find, collapse, and reverse-apply logic.

function Add-BatesHyperlinksToRange
{
    param(
        $Document,
        $SearchRange,
        [string]$SearchText,
        [string]$TargetPath,
        [int]$SafetyLimit = 50
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
            Write-Host "        Safety limit reached while searching for $SearchText" -ForegroundColor Yellow
            Write-Host "            Not all occurences of the above have been hyperlinked."
            Write-Host "            Consider increaseing SafetyLimit parameter, currently " -NoNewline
            Write-Host $SafetyLimit -ForegroundColor Yellow
            break
        }

        # Move past the current match before searching again.

        $SearchRange.Collapse($wdCollapseEnd)
    }

    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Find) | Out-Null

    # Locate all matches first, then apply hyperlinks from last to first.
    # Applying links during the search could make Find match the Bates ID inside
    # the newly inserted hyperlink field address.

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



Write-Host "Checking for source folders..."

# Find matching folders in the current directory.
$EvidenceFolders = Get-ChildItem -Directory |
    Where-Object { $_.Name -cIn $SupportedFolders }

# Validate the source folder selection.
if ($EvidenceFolders.Count -eq 0) {
    Write-Error "No 'Evidence' or 'Documents' folder found."
    return
}

if ($EvidenceFolders.Count -gt 1) {
    Write-Error "Both 'Evidence' and 'Documents' folders were found. Only one may exist."
    return
}

$EvidenceFolder = $EvidenceFolders[0]

Write-Host "Using evidence source folder:" -NoNewline
Write-Host " $($EvidenceFolder.Name)" -ForegroundColor Green

# --- Enumerate source files ---

$EnumerationStart = Get-Date

$EvidenceFiles = Get-ChildItem -Path $EvidenceFolder.FullName -File

Write-Host "Found $($EvidenceFiles.Count) files."

if ($EvidenceFiles.Count -eq 0)
{
    Write-Error "The '$($EvidenceFolder.Name)' folder is empty. Cannot continue."
    return
}

# --- Build evidence lookup ---

$EvidenceLookup = @{}

$FileCounter = 0
$CalloutInterval = 5000

$InvalidEvidenceFiles = @()

foreach ($File in $EvidenceFiles) {

    $FileCounter++

    if ($FileCounter % $CalloutInterval -eq 0) {
        Write-Host "    $FileCounter files indexed."
    }

    # Use the filename without its extension as the evidence ID.

    $EvidenceID = $File.BaseName.Trim().ToUpper()

    # Validate the filename format; an optional -N or -NN suffix is supported.

    if ($EvidenceID -notmatch '^[A-Z]{3}\.\d{3}\.\d{3}\.\d{3,4}(-\d{1,2})?$') {

        $InvalidEvidenceFiles += $File.Name

        continue
    }

    # Reject duplicate evidence IDs.

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
}

# --- Report invalid filenames ---

if ($InvalidEvidenceFiles.Count -gt 0) {

    Write-Host ""
    Write-Host "Files ignored because their filenames do not match the expected format: " -NoNewline
    Write-Host $InvalidEvidenceFiles.Count -ForegroundColor Yellow

    foreach ($InvalidFile in ($InvalidEvidenceFiles | Sort-Object)) {
        Write-Host "    $InvalidFile" -ForegroundColor Yellow
    }
    Write-Host "These files will be ignored; valid files will continue to be indexed."
}

$EnumerationDuration = (Get-Date) - $EnumerationStart

Write-Host ""
Write-Host "File indexing complete."
Write-Host "    Indexed files: " -NoNewline
Write-Host $EvidenceLookup.Count -ForegroundColor Green
Write-Host "    Duration: " -NoNewline
Write-Host $EnumerationDuration.ToString('hh\:mm\:ss') -ForegroundColor Green

# --- Enumerate Word documents ---

if (-not (Test-Path -LiteralPath $TargetDir -PathType Container)) {
    Write-Error "The '$TargetDir' folder does not exist."
    return
}

$WordDocuments = Get-ChildItem -Path $TargetDir -File -Filter "*.docx" |
    Where-Object { $_.Name -notlike "~$*" } |
    Sort-Object Name

if ($WordDocuments.Count -eq 0) {
    Write-Error "No .docx files found in $TargetDir."
    return
}

# --- Process Word documents ---
# Looping allows multiple documents to be handled in a single run.

# Accumulated across every document processed this run, for the final errata files.
$AllReferencedBates = @{}
$AllMissingBates = @{}

$RemainingDocuments = New-Object System.Collections.ArrayList
foreach ($WordDocument in $WordDocuments) { [void]$RemainingDocuments.Add($WordDocument) }

while ($RemainingDocuments.Count -gt 0)
{

# Reset per document so the reported total runtime covers only this document.
$ScriptStart = Get-Date

# --- Select the Word document ---
# Always prompt so the user confirms the document, even when only one remains.

Write-Host ""
Write-Host "Available Word documents:"
Write-Host ""

for ($i = 0; $i -lt $RemainingDocuments.Count; $i++)
{
    Write-Host "[$($i + 1)] $($RemainingDocuments[$i].Name)"
}

Write-Host ""

$Selection = Read-Host "Enter document number (blank to exit)"

$ValidSelection = (
    $Selection -match '^\d+$' -and
    [int]$Selection -ge 1 -and
    [int]$Selection -le $RemainingDocuments.Count
)

if (-not $ValidSelection)
{
    Write-Host ""
    Write-Host "No document selected. Exiting." -ForegroundColor Yellow
    break
}

$SelectedDocument = $RemainingDocuments[[int]$Selection - 1]

Write-Host ""
Write-Host "Selected document:" -NoNewline
Write-Host " $($SelectedDocument.Name)" -ForegroundColor Green

$WordDocumentPath = $SelectedDocument.FullName

# --- Open the Word document ---

Write-Host ""
Write-Host "Opening Word document..."

$Word = $null
$Document = $null
$DocumentWasOpened = $false
$TotalLinksAdded = 0

try
{
    $Word = New-Object -ComObject Word.Application
    $Word.Visible = $false
    $Word.DisplayAlerts = 0

    $Document = $Word.Documents.Open($WordDocumentPath)
    $DocumentWasOpened = $true

    Write-Host "Document opened successfully." -ForegroundColor Green

# --- Extract document text ---
# Only the main body and footnotes are scanned and linked. Headers, footers,
# endnotes, comments, and text boxes are excluded.

Write-Host ""
Write-Host "Extracting document text..."

# Main document body.

$BodyText = $Document.Content.Text

# Footnotes.

$FootnoteTextBuilder = New-Object System.Text.StringBuilder

foreach ($Footnote in $Document.Footnotes)
{
    if ($Footnote.Range.Text)
    {
        $null = $FootnoteTextBuilder.AppendLine($Footnote.Range.Text)
    }
}

$FootnoteText = $FootnoteTextBuilder.ToString()

# --- Document statistics ---

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

# --- Scan the main document body ---

$ScanStart = Get-Date

# Reject adjacent letters or digits to avoid partial matches, while allowing
# punctuation such as ",./)". An optional -N or -NN suffix is supported.
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

# --- Scan footnotes ---

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

# --- Combine Bates references ---

$AllBatesLookup = Merge-BatesLookup -Lookups @($BodyBatesLookup, $FootnoteBatesLookup)

Write-Host ""
Write-Host "Document Bates summary."

Write-Host "    Total unique Bates references: " -NoNewline
Write-Host $AllBatesLookup.Count -ForegroundColor Green

if ($AllBatesLookup.Count -eq 0)
{
    throw "No Bates references found in the Word document."
}

# --- Reconcile references with evidence files ---

$Reconciliation = Get-BatesReconciliation -DocumentBatesLookup $AllBatesLookup -EvidenceLookup $EvidenceLookup

$MatchedCount = $Reconciliation.MatchedCount
$MissingFromFolder = $Reconciliation.MissingFromFolder
$UnreferencedEvidence = $Reconciliation.UnreferencedEvidence

$UnreferencedFiles = foreach ($EvidenceID in ($UnreferencedEvidence.Keys | Sort-Object))
{
    $EvidenceLookup[$EvidenceID]
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
        Write-Host "        $Bates" -ForegroundColor Yellow
    }
}

Write-Host "    Evidence not referenced (file in folder but not referred to in document): " -NoNewline
Write-Host $UnreferencedEvidence.Count -ForegroundColor Yellow

if ($UnreferencedEvidence.Count -gt 0)
{
    foreach ($Bates in ($UnreferencedEvidence.Keys | Sort-Object))
    {
        Write-Host "        $Bates" -ForegroundColor Yellow
    }
}
# --- Hyperlink matched Bates references ---
# Search only the main body and footnote stories, matching the scan scope above.
# Word's Find engine is used because Hyperlinks.Add requires a live Range rather
# than a plain string offset.

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
        $TargetPath = "../" + $EvidenceFolder.Name + "/" + $EvidenceLookup[$CurrentBates]

        # Main document body.

        $BodyRange = $Document.Content.Duplicate

        $BodyMatchCount = Add-BatesHyperlinksToRange -Document $Document -SearchRange $BodyRange -SearchText $CurrentBates -TargetPath $TargetPath

        $BodyLinksAdded += $BodyMatchCount

        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($BodyRange) | Out-Null

        # Footnotes.

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


# --- Handle errors and close Word ---
}
catch
{
    Write-Host "Could not finish processing this document." -ForegroundColor Red
    Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
}
finally
{
    Write-Host ""

	if ($Document)
	{
		$TotalLinksAdded = $BodyLinksAdded + $FootnoteLinksAdded

        # Save only when at least one hyperlink was created.
		if ($TotalLinksAdded -ge 1)
		{
			Write-Host "Saving document..."

            # Clear the undo history created by hyperlink insertion before saving.
			$Document.UndoClear()

			$Document.Save()

			Write-Host "Document saved." -ForegroundColor Green
		}
		else
		{
			Write-Host "Document not saved because no hyperlinks were created." -ForegroundColor Yellow
		}
	}

    if ($Document)
    {
        Write-Host ""
        Write-Host "Closing Word document..."

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

    if ($DocumentWasOpened)
    {
        Write-Host "Word closed." -ForegroundColor Green
    }

    # Merge this document's reconciliation into the run-wide totals.
    if ($TotalLinksAdded -gt 1)
    {
        foreach ($Bates in $MatchedBates) { $AllReferencedBates[$Bates] = $true }
        foreach ($Bates in $MissingFromFolder.Keys) { $AllMissingBates[$Bates] = $true }
    }

    $TotalRuntime = (Get-Date) - $ScriptStart
    Write-Host "Total runtime: " -NoNewline
    Write-Host $TotalRuntime.ToString('hh\:mm\:ss') -ForegroundColor Green

    Write-ExecutionLog -FolderPath $LogFolder -HyperlinksCreated $TotalLinksAdded
}

$RemainingDocuments.Remove($SelectedDocument)

if ($RemainingDocuments.Count -eq 0)
{
    break
}

}

# --- Write errata files ---
# Deduped across every document processed this run: a Bates ID only ends up in
# _no_reference.txt if no processed document referenced it, and only ends up in
# _no_files.txt if it was referenced but never matched to an evidence file.

if ($AllReferencedBates.Count -gt 0 -or $AllMissingBates.Count -gt 0)
{
    $FinalUnreferencedFiles = foreach ($EvidenceID in ($EvidenceLookup.Keys | Sort-Object))
    {
        if (-not $AllReferencedBates.ContainsKey($EvidenceID))
        {
            $EvidenceLookup[$EvidenceID]
        }
    }

    $InvalidEvidenceCount = $InvalidEvidenceFiles.Count
    $UnreferencedCount = $FinalUnreferencedFiles.Count
    $MissingCount = $AllMissingBates.Count

    Write-ListingFile -Path $NoBatesPath -Items ($InvalidEvidenceFiles | Sort-Object)
    Write-ListingFile -Path $NoReferencePath -Items $FinalUnreferencedFiles
    Write-ListingFile -Path $NoFilesPath -Items ($AllMissingBates.Keys | Sort-Object)

    Write-Host ""
    Write-Host "Errata files written."

    Write-Host "    $(Split-Path -Path $NoBatesPath -Leaf) : " -NoNewline
    Write-Host $InvalidEvidenceCount -ForegroundColor Yellow -NoNewline
    Write-Host " item(s)"

    Write-Host "    $(Split-Path -Path $NoReferencePath -Leaf) : " -NoNewline
    Write-Host $UnreferencedCount -ForegroundColor Yellow -NoNewline
    Write-Host " item(s)"

    Write-Host "    $(Split-Path -Path $NoFilesPath -Leaf) : " -NoNewline
    Write-Host $MissingCount -ForegroundColor Yellow -NoNewline
    Write-Host " item(s)"
    Write-Host ""
}
else
{
    Write-Host ""
    Write-Host "No errata files were created." -ForegroundColor Yellow
}