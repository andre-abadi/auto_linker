# Excel Document ID Auto-Linker

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
