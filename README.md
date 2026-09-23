# Part 1: Word Document Auto-Linker

`auto_linker_word.ps1` scans Word documents for Bates document IDs and creates hyperlinks to the matching evidence files. It is intended for a working directory containing:

- A `Covering Material` folder containing one or more `.docx` files.
- Exactly one document source folder named `Evidence` or `Documents`.
- Evidence filenames whose base names match Bates IDs, such as `AAA.111.222.333.txt` or `AAA.111.222.4444.txt`.

The script checks the main document body and footnotes for Bates IDs in the format `AAA.111.222.333` *(TLA.3.3.3)* or `AAA.111.222.333` *(TLA.3.3.4)*. An optional suffix such as `-1` or `-12` is supported. It reports invalid filenames, reconciles document references with the evidence folder, and creates hyperlinks for every matching Bates ID. Word documents are saved only when at least one hyperlink is added.

## Word prerequisites

- Windows with PowerShell.
- Microsoft Word installed, because the script uses Word COM automation.
- The required folders and files arranged as described above.

## Run the Word script as a PowerShell script

Open PowerShell in the project directory and run:

 ```powershell
 .\auto_linker_word.ps1
 ```

Select each document when prompted. Enter a blank response to stop selecting documents. At the end, enter `Y` when prompted if you want the script to write the errata files.

## Run the Word script as a text file

If PowerShell will not execute the file as a `.ps1`, make a copy named `auto_linker_word.txt` and run its contents with `Invoke-Expression` from the same directory:

 ```powershell
 Invoke-Expression (Get-Content .\auto_linker_word.txt -Raw)
 ```

Running the contents as text still requires PowerShell and Microsoft Word, and the command must be run from the directory containing `Covering Material` and the evidence folder.

## Word outputs

- Matching Bates IDs in the selected Word documents become hyperlinks to the evidence files.
- Optionally:
   - `_no_bates.txt` lists evidence files with invalid Bates filenames.
   - `_no_reference.txt` lists evidence files that were not referenced by any processed document.
   - `_no_files.txt` lists Bates IDs referenced by a document but missing from the evidence folder.

## Word performance and complexity

At a high level, the script has three main costs:

- Evidence indexing is approximately linear in the number of evidence files, `O(E)`.
- Bates scanning is approximately linear in the amount of text in each Word document, `O(T)`.
- Hyperlinking is the main bottleneck. For each distinct matched Bates ID, Word searches the document body and footnotes again. In the worst case this behaves like `O(B x T)`, where `B` is the number of distinct matched Bates IDs and `T` is the document text size.

This repeated search can look quadratic when the number of Bates IDs and the document size grow together. The dominant practical cost is Word COM automation and inserting hyperlinks, not the PowerShell hashtables used for lookups. Multiple large documents multiply these costs across the run.


# Part 2: Excel Document ID Auto-Linker

A PowerShell automation script that creates hyperlinks between Document IDs in Excel worksheets and their corresponding files in a directory.

## Purpose

This script streamlines document management by automatically creating hyperlinks in Excel spreadsheets that contain Document ID references. It's particularly useful for organizations that maintain large document libraries and need efficient ways to link between document indexes and actual files.

## Features

- Processes Excel files containing Document ID columns
- Automatically detects Document ID columns in the first 3 rows of each worksheet
- Creates hyperlinks between Document IDs and their corresponding files
- Supports multiple worksheets in a single Excel file
- Generates detailed progress reports during execution
- Creates an errata file listing any missing or extra documents
- Provides execution time reporting
- Includes error handling and COM object cleanup
- Generates execution logs for troubleshooting

## Performance

- Includes progress indicators for large file sets
- Implements garbage collection for memory management
- Optimized for handling large document libraries
- Tested on 110,000 files and the same number of hyperlinks.
  - Staying under the hard coded maximum of [65,530 hyperlinks per worksheet](https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3).
  - Total script runtime was approximately 10mins.

## Prerequisites

- Windows environment with PowerShell
- Microsoft Excel installed
- Single Excel file (.xlsx) in the working directory
- Single folder containing the documents to be linked
- Document IDs in Excel must match filenames (without extensions)

## Usage

1. Place the script in a directory containing:
   - One Excel file (.xlsx)
   - One folder containing the documents to be linked
2. Run the script in PowerShell
3. The script will:
   - Process all visible worksheets
   - Create hyperlinks for matching Document IDs
   - Generate an errata file if any documents are missing or unmatched
   - Create a log file of the execution

## Output

- Updated Excel file with hyperlinks
- Errata file (if needed) listing missing or extra documents
- Execution log file with detailed processing information

## Notes

- The script searches for "Document ID" column headers (configurable)
- Hidden worksheets are automatically skipped
- Relative paths are used for hyperlinks
- Case-insensitive document ID matching
