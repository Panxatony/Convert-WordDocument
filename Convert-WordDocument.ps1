<#
.Synopsis
PowerShell script to convert Word documents - Improved Version

.Description
This script converts Word compatible documents to a selected format utilizing the Word SaveAs function.
Improved version with better error handling, COM cleanup, and performance optimizations.

The script converts either all documents in a single folder matching an include filter or a single file.

Currently supported target document types:
- Default --> Word 2016 (DOCX)
- PDF
- XPS
- HTML
- RTF

Author: Thomas Stensitzki (Original)
Improved by: Claude AI Assistant

Version 2.0 2025-05-29

.NOTES 
Requirements 
- Word 2016+ installed locally
- PowerShell 5.1 or later

Revision History 
-------------------------------------------------------------------------------- 
1.0      Initial release
1.1      Updated Word cleanup code
2.0      Major improvements: Better COM cleanup, input validation, progress bars,
         performance optimizations, enhanced error handling

.LINK
http://scripts.granikos.eu

.PARAMETER SourcePath
Source path to a folder containing the documents to convert or full path to a single document

.PARAMETER IncludeFilter
File extension filter when converting all files in a single folder. Default: *.doc

.PARAMETER TargetFormat
Word Save AS target format. Currently supported: Default, PDF, XPS, HTML, RTF

.PARAMETER DeleteExistingFiles
Switch to delete an existing target file

.PARAMETER ReuseWordInstance
Switch to reuse a single Word instance for multiple files (better performance)

.PARAMETER Quiet
Switch to suppress non-error output

.EXAMPLE
Convert all .doc files in E:\temp to Default format

.\Convert-WordDocument.ps1 -SourcePath E:\Temp -IncludeFilter *.doc 

.EXAMPLE
Convert all .doc files in E:\temp to PDF with progress display

.\Convert-WordDocument.ps1 -SourcePath E:\Temp -IncludeFilter *.doc -TargetFormat PDF

.EXAMPLE
Convert a single document to Word default format

.\Convert-WordDocument.ps1 -SourcePath E:\Temp\MyDocument.doc

.EXAMPLE
Convert multiple files efficiently by reusing Word instance

.\Convert-WordDocument.ps1 -SourcePath E:\Temp -TargetFormat PDF -ReuseWordInstance
#>

[CmdletBinding()]
Param(
    [Parameter(Mandatory=$true, HelpMessage="Source path to folder or single file")]
    [ValidateScript({
        if (Test-Path $_) { $true }
        else { throw "Path '$_' does not exist" }
    })]
    [string]$SourcePath,
    
    [ValidateNotNullOrEmpty()]
    [string]$IncludeFilter = '*.doc',
    
    [ValidateSet('Default','PDF','XPS','HTML','RTF')]
    [string]$TargetFormat = 'Default',
    
    [switch]$DeleteExistingFiles,
    
    [switch]$ReuseWordInstance,
    
    [switch]$Quiet
)

# Error codes
$ERR_OK = 0
$ERR_COMOBJECT = 1001 
$ERR_SOURCEPATHMISSING = 1002
$ERR_WORDSAVEAS = 1003
$ERR_WORDNOTINSTALLED = 1004
$ERR_INVALIDFILE = 1005

# Define Word target document types
$wdFormat = @{
    'Document' = 0
    'Template' = 1
    'RTF' = 6
    'HTML' = 8
    'Default' = 16
    'PDF' = 17
    'XPS' = 18
}

$FileExtension = @{
    'Document' = '.doc'
    'Template' = '.dot'
    'RTF' = '.rtf'
    'HTML' = '.html'
    'Default' = '.docx'
    'PDF' = '.pdf'
    'XPS' = '.xps'
}

# Supported file extensions for input
$SupportedExtensions = @('.doc', '.docx', '.dot', '.dotx', '.rtf')

function Test-WordInstallation {
    <#
    .SYNOPSIS
    Tests if Microsoft Word is installed and accessible
    #>
    try {
        $testWord = New-Object -ComObject Word.Application -ErrorAction Stop
        $testWord.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($testWord) | Out-Null
        [GC]::Collect()
        return $true
    }
    catch {
        return $false
    }
}

function Write-LogMessage {
    <#
    .SYNOPSIS
    Writes log messages with proper formatting
    #>
    param(
        [string]$Message,
        [ValidateSet('Info','Warning','Error','Verbose')]
        [string]$Level = 'Info'
    )
    
    if ($Quiet -and $Level -eq 'Info') { return }
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    switch ($Level) {
        'Info' { Write-Host "[$timestamp] INFO: $Message" -ForegroundColor Green }
        'Warning' { Write-Warning "[$timestamp] WARNING: $Message" }
        'Error' { Write-Error "[$timestamp] ERROR: $Message" }
        'Verbose' { Write-Verbose "[$timestamp] VERBOSE: $Message" }
    }
}

function Invoke-SafeComCleanup {
    <#
    .SYNOPSIS
    Safely cleans up COM objects
    #>
    param(
        [object]$WordDocument,
        [object]$WordApplication
    )
    
    try {
        if ($WordDocument -ne $null) {
            try {
                $WordDocument.Close([ref]$false)
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($WordDocument) | Out-Null
            }
            catch {
                Write-LogMessage "Warning during document cleanup: $($_.Exception.Message)" -Level Warning
            }
        }
        
        if ($WordApplication -ne $null) {
            try {
                $WordApplication.Quit()
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($WordApplication) | Out-Null
            }
            catch {
                Write-LogMessage "Warning during application cleanup: $($_.Exception.Message)" -Level Warning
            }
        }
    }
    catch {
        Write-LogMessage "Error during COM cleanup: $($_.Exception.Message)" -Level Error
    }
    finally {
        # Force garbage collection
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()
        
        # Remove variables if they exist
        if (Get-Variable -Name WordDocument -ErrorAction SilentlyContinue) {
            Remove-Variable -Name WordDocument -Force -ErrorAction SilentlyContinue
        }
        if (Get-Variable -Name WordApplication -ErrorAction SilentlyContinue) {
            Remove-Variable -Name WordApplication -Force -ErrorAction SilentlyContinue
        }
    }
}

function ConvertTo-WordDocument {
    <#
    .SYNOPSIS
    Converts a single Word document to the specified format
    #>
    [CmdletBinding()]
    Param(
        [string]$FileSourcePath,
        [string]$SourceFileExtension,
        [string]$TargetFileExtension,
        [int]$WdSaveFormat = 16,
        [switch]$DeleteFile,
        [object]$ExistingWordApp = $null
    )

    $WordApplication = $null
    $WordDocument = $null
    $shouldCleanupApp = $false

    try {
        # Validate input file
        if (-not (Test-Path -Path $FileSourcePath)) {
            throw "Source file not found: $FileSourcePath"
        }

        # Check if file extension is supported
        $fileInfo = Get-Item -Path $FileSourcePath
        if ($fileInfo.Extension.ToLower() -notin $SupportedExtensions) {
            Write-LogMessage "Skipping unsupported file: $($fileInfo.Name)" -Level Warning
            return $false
        }

        Write-LogMessage "Converting: $($fileInfo.Name) -> $TargetFormat"

        # Use existing Word application or create new one
        if ($ExistingWordApp -ne $null) {
            $WordApplication = $ExistingWordApp
        }
        else {
            try {
                $WordApplication = New-Object -ComObject Word.Application
                $WordApplication.Visible = $false
                $WordApplication.DisplayAlerts = 0  # Disable alerts
                $shouldCleanupApp = $true
            }
            catch {
                throw "Could not create Word COM object: $($_.Exception.Message)"
            }
        }

        # Open document
        try {
            $WordDocument = $WordApplication.Documents.Open($FileSourcePath, $false, $true) # ReadOnly = true
        }
        catch {
            throw "Could not open document '$FileSourcePath': $($_.Exception.Message)"
        }

        # Generate target file path
        $NewFilePath = ($FileSourcePath).Replace($SourceFileExtension, $TargetFileExtension)

        # Handle existing files
        if (Test-Path -Path $NewFilePath) {
            if ($DeleteFile) {
                try {
                    Remove-Item -Path $NewFilePath -Force -Confirm:$false
                    Write-LogMessage "Deleted existing file: $NewFilePath" -Level Verbose
                }
                catch {
                    throw "Could not delete existing file '$NewFilePath': $($_.Exception.Message)"
                }
            }
            else {
                Write-LogMessage "Target file already exists: $NewFilePath (use -DeleteExistingFiles to overwrite)" -Level Warning
                return $false
            }
        }

        # Save document in new format
        try {
            $WordDocument.SaveAs([ref]$NewFilePath, [ref]$WdSaveFormat)
            Write-LogMessage "Successfully converted to: $NewFilePath"
            return $true
        }
        catch {
            throw "Could not save document as '$NewFilePath': $($_.Exception.Message)"
        }
    }
    catch {
        Write-LogMessage "Error converting '$FileSourcePath': $($_.Exception.Message)" -Level Error
        return $false
    }
    finally {
        # Clean up document (but not application if reusing)
        if ($WordDocument -ne $null) {
            try {
                $WordDocument.Close([ref]$false)
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($WordDocument) | Out-Null
            }
            catch {
                Write-LogMessage "Warning during document cleanup: $($_.Exception.Message)" -Level Warning
            }
        }

        # Only cleanup application if we created it (not reusing)
        if ($shouldCleanupApp -and $WordApplication -ne $null) {
            Invoke-SafeComCleanup -WordDocument $null -WordApplication $WordApplication
        }

        [GC]::Collect()
    }
}

# Main execution
try {
    Write-LogMessage "Starting Word document conversion..."
    Write-LogMessage "Source: $SourcePath"
    Write-LogMessage "Target Format: $TargetFormat"

    # Test Word installation
    if (-not (Test-WordInstallation)) {
        Write-LogMessage "Microsoft Word is not installed or not accessible" -Level Error
        exit $ERR_WORDNOTINSTALLED
    }

    # Determine if source is folder or file
    $IsFolder = (Get-Item -Path $SourcePath) -is [System.IO.DirectoryInfo]

    if ($IsFolder) {
        # Process folder
        Write-LogMessage "Processing folder: $SourcePath"
        
        $SourceFiles = Get-ChildItem -Path $SourcePath -Include $IncludeFilter -Recurse -File
        $totalFiles = ($SourceFiles | Measure-Object).Count

        if ($totalFiles -eq 0) {
            Write-LogMessage "No files found matching filter '$IncludeFilter' in '$SourcePath'" -Level Warning
            exit $ERR_OK
        }

        Write-LogMessage "Found $totalFiles files to convert"

        $WordApplication = $null
        $convertedCount = 0
        $failedCount = 0

        try {
            # Create Word application if reusing instance
            if ($ReuseWordInstance) {
                Write-LogMessage "Creating reusable Word instance for better performance..."
                $WordApplication = New-Object -ComObject Word.Application
                $WordApplication.Visible = $false
                $WordApplication.DisplayAlerts = 0
            }

            # Process each file
            $currentFile = 0
            foreach ($File in $SourceFiles) {
                $currentFile++
                
                # Show progress
                if (-not $Quiet) {
                    Write-Progress -Activity "Converting Documents" -Status "Processing $($File.Name)" -PercentComplete (($currentFile / $totalFiles) * 100)
                }

                $success = ConvertTo-WordDocument -FileSourcePath $File.FullName -SourceFileExtension $File.Extension -TargetFileExtension $FileExtension.Item($TargetFormat) -WdSaveFormat $wdFormat.Item($TargetFormat) -DeleteFile:$DeleteExistingFiles -ExistingWordApp $WordApplication

                if ($success) {
                    $convertedCount++
                }
                else {
                    $failedCount++
                }
            }

            # Complete progress bar
            if (-not $Quiet) {
                Write-Progress -Activity "Converting Documents" -Completed
            }
        }
        finally {
            # Cleanup reused Word instance
            if ($WordApplication -ne $null) {
                Invoke-SafeComCleanup -WordDocument $null -WordApplication $WordApplication
            }
        }

        Write-LogMessage "Conversion completed. Successful: $convertedCount, Failed: $failedCount"
    }
    else {
        # Process single file
        Write-LogMessage "Processing single file: $SourcePath"
        
        $File = Get-Item -Path $SourcePath
        $success = ConvertTo-WordDocument -FileSourcePath $File.FullName -SourceFileExtension $File.Extension -TargetFileExtension $FileExtension.Item($TargetFormat) -WdSaveFormat $wdFormat.Item($TargetFormat) -DeleteFile:$DeleteExistingFiles

        if ($success) {
            Write-LogMessage "Single file conversion completed successfully"
        }
        else {
            Write-LogMessage "Single file conversion failed" -Level Error
            exit $ERR_WORDSAVEAS
        }
    }

    Write-LogMessage "Script execution completed successfully"
    exit $ERR_OK
}
catch {
    Write-LogMessage "Critical error: $($_.Exception.Message)" -Level Error
    exit $ERR_WORDSAVEAS
}
finally {
    # Final cleanup
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
