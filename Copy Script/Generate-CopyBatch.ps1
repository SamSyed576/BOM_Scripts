<#
.SYNOPSIS
    Processes multiple SolidWorks BOM CSV files and generates a batch file
    to copy all unique ParentPath and ChildPath files.

.DESCRIPTION
    Features:
    - Multiple CSV files
    - Dynamic delimiter detection
    - Supports single and multi-character delimiters
    - Examples: ',', '|', '##', '@@', '||', etc.
    - Case-insensitive ParentPath / ChildPath detection
    - Handles quoted fields
    - Ignores additional columns
    - Removes duplicate file paths
    - Processes one CSV at a time
    - XCOPY or ROBOCOPY
    - Preserves folder structure
    - Missing file report
    - Copy log
    - Optional AutoRun
    - No PAUSE command in generated BAT

.USAGE

    powershell -ExecutionPolicy Bypass `
        -File Generate-CopyBatch.ps1

    powershell -ExecutionPolicy Bypass `
        -File Generate-CopyBatch.ps1 `
        -ConfigPath "D:\tools\config.ini"
#>

param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.ini")
)

# ============================================================
# READ INI FILE
# ============================================================

function Read-IniFile {
    param(
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Config file not found: $Path"
    }

    $ini = @{}
    $section = "NoSection"
    $ini[$section] = @{}

    foreach ($line in Get-Content -LiteralPath $Path) {

        $trimmed = $line.Trim()

        if (
            [string]::IsNullOrWhiteSpace($trimmed) -or
            $trimmed.StartsWith(";") -or
            $trimmed.StartsWith("#")
        ) {
            continue
        }

        if ($trimmed -match '^\[(.+)\]$') {

            $section = $matches[1].Trim()

            if (-not $ini.ContainsKey($section)) {
                $ini[$section] = @{}
            }

            continue
        }

        if ($trimmed -match '^([^=]+)=(.*)$') {

            $key = $matches[1].Trim()
            $value = $matches[2].Trim()

            $ini[$section][$key] = $value
        }
    }

    return $ini
}

# ============================================================
# CONFIG VALUE
# ============================================================

function Get-ConfigValue {
    param(
        [hashtable]$Config,
        [string]$Section,
        [string]$Key
    )

    if (
        -not $Config.ContainsKey($Section) -or
        -not $Config[$Section].ContainsKey($Key)
    ) {
        return $null
    }

    return $Config[$Section][$Key]
}

# ============================================================
# REMOVE QUOTES
# ============================================================

function Remove-OuterQuotes {
    param(
        [string]$Value
    )

    if ($null -eq $Value) {
        return ""
    }

    $Value = $Value.Trim()

    if (
        $Value.Length -ge 2 -and
        $Value.StartsWith('"') -and
        $Value.EndsWith('"')
    ) {
        $Value = $Value.Substring(
            1,
            $Value.Length - 2
        )
    }

    return $Value.Trim()
}

# ============================================================
# PARSE DELIMITED LINE
#
# Handles:
#   comma
#   pipe
#   ##
#   @@
#   ||
#   etc.
#
# Also handles quoted fields.
#
# Example:
#
# "Assembly##01.SLDASM"##"Part##01.SLDPRT"
#
# The ## inside quotes is NOT treated as a delimiter.
# ============================================================

function Parse-DelimitedLine {
    param(
        [string]$Line,
        [string]$Delimiter
    )

    $fields = New-Object System.Collections.Generic.List[string]

    $current = New-Object System.Text.StringBuilder

    $insideQuotes = $false

    $i = 0

    while ($i -lt $Line.Length) {

        $char = $Line[$i]

        # ----------------------------------------------------
        # Handle quote
        # ----------------------------------------------------

        if ($char -eq '"') {

            # Escaped quote inside quoted field: ""
            if (
                $insideQuotes -and
                ($i + 1 -lt $Line.Length) -and
                $Line[$i + 1] -eq '"'
            ) {

                [void]$current.Append('"')

                $i += 2
                continue
            }

            $insideQuotes = -not $insideQuotes

            $i++
            continue
        }

        # ----------------------------------------------------
        # Check delimiter only outside quotes
        # ----------------------------------------------------

        if (-not $insideQuotes) {

            if (
                $Delimiter.Length -gt 0 -and
                ($i + $Delimiter.Length -le $Line.Length)
            ) {

                $possibleDelimiter =
                    $Line.Substring(
                        $i,
                        $Delimiter.Length
                    )

                if ($possibleDelimiter -eq $Delimiter) {

                    $fields.Add(
                        $current.ToString().Trim()
                    )

                    $current.Clear()

                    $i += $Delimiter.Length
                    continue
                }
            }
        }

        [void]$current.Append($char)

        $i++
    }

    # Add final field
    $fields.Add(
        $current.ToString().Trim()
    )

    return $fields
}

# ============================================================
# DETECT DELIMITER
#
# We test candidate delimiters against the header.
#
# A delimiter is accepted only if BOTH configured columns
# are found after splitting.
# ============================================================

function Find-Delimiter {
    param(
        [string]$Header,
        [string]$ParentColumn,
        [string]$ChildColumn
    )

    # Longer delimiters first.
    #
    # This prevents:
    #
    # ## being interpreted as two #
    #
    $candidateDelimiters = @(
        "##",
        "@@",
        "||",
        ";;",
        "::",
        "<->",
        "|||",
        ",",
        "|",
        ";",
        "`t"
    )

    foreach ($delimiter in $candidateDelimiters) {

        $columns = Parse-DelimitedLine `
            -Line $Header `
            -Delimiter $delimiter

        $parentFound = $false
        $childFound = $false

        foreach ($column in $columns) {

            $cleanColumn = Remove-OuterQuotes $column

            if (
                $cleanColumn.Equals(
                    $ParentColumn,
                    [System.StringComparison]::OrdinalIgnoreCase
                )
            ) {
                $parentFound = $true
            }

            if (
                $cleanColumn.Equals(
                    $ChildColumn,
                    [System.StringComparison]::OrdinalIgnoreCase
                )
            ) {
                $childFound = $true
            }
        }

        if ($parentFound -and $childFound) {
            return $delimiter
        }
    }

    return $null
}

# ============================================================
# NORMALIZE FILE PATH
# ============================================================

function Normalize-FilePath {
    param(
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $null
    }

    $Path = Remove-OuterQuotes $Path

    # Convert / to \
    $Path = $Path.Replace("/", "\")

    return $Path.Trim()
}

# ============================================================
# LOAD CONFIG
# ============================================================

Write-Host ""
Write-Host "============================================================"
Write-Host " SolidWorks BOM Copy Batch Generator"
Write-Host "============================================================"
Write-Host ""

Write-Host "Loading configuration:"
Write-Host "$ConfigPath"
Write-Host ""

$config = Read-IniFile -Path $ConfigPath

# ============================================================
# GET CONFIGURATION
# ============================================================

$reportFolder = Get-ConfigValue `
    $config `
    "Paths" `
    "BOMReportFolder"

$destPath = Get-ConfigValue `
    $config `
    "Paths" `
    "DestinationPath"

$parentCol = Get-ConfigValue `
    $config `
    "BOM" `
    "ParentPathColumn"

$childCol = Get-ConfigValue `
    $config `
    "BOM" `
    "ChildPathColumn"

$copyTool = Get-ConfigValue `
    $config `
    "Settings" `
    "CopyTool"

$preserveTreeValue = Get-ConfigValue `
    $config `
    "Settings" `
    "PreserveFolderStructure"

$autoRunValue = Get-ConfigValue `
    $config `
    "Settings" `
    "AutoRun"

$xcopySw = Get-ConfigValue `
    $config `
    "XCOPY" `
    "Switches"

$robocopySw = Get-ConfigValue `
    $config `
    "ROBOCOPY" `
    "Switches"

# ============================================================
# VALIDATE CONFIG
# ============================================================

if ([string]::IsNullOrWhiteSpace($reportFolder)) {
    throw "BOMReportFolder is missing from config.ini"
}

if ([string]::IsNullOrWhiteSpace($destPath)) {
    throw "DestinationPath is missing from config.ini"
}

if ([string]::IsNullOrWhiteSpace($parentCol)) {
    throw "ParentPathColumn is missing from config.ini"
}

if ([string]::IsNullOrWhiteSpace($childCol)) {
    throw "ChildPathColumn is missing from config.ini"
}

if ([string]::IsNullOrWhiteSpace($copyTool)) {
    throw "CopyTool is missing from config.ini"
}

$copyTool = $copyTool.Trim().ToUpper()

if (
    $copyTool -ne "XCOPY" -and
    $copyTool -ne "ROBOCOPY"
) {
    throw "CopyTool must be XCOPY or ROBOCOPY"
}

$preserveTree = (
    $preserveTreeValue.Trim().Equals(
        "true",
        [System.StringComparison]::OrdinalIgnoreCase
    )
)

$autoRun = (
    $autoRunValue.Trim().Equals(
        "true",
        [System.StringComparison]::OrdinalIgnoreCase
    )
)

# ============================================================
# VALIDATE INPUT FOLDER
# ============================================================

if (-not (Test-Path -LiteralPath $reportFolder)) {
    throw "BOM folder not found: $reportFolder"
}

# ============================================================
# CREATE DESTINATION
# ============================================================

if (-not (Test-Path -LiteralPath $destPath)) {

    Write-Host "Creating destination:"
    Write-Host "$destPath"

    New-Item `
        -ItemType Directory `
        -Path $destPath `
        -Force |
        Out-Null
}

# ============================================================
# DISPLAY CONFIG
# ============================================================

Write-Host ""
Write-Host "Configuration"
Write-Host "------------------------------------------------------------"
Write-Host "BOM folder          : $reportFolder"
Write-Host "Destination         : $destPath"
Write-Host "Parent column       : $parentCol"
Write-Host "Child column        : $childCol"
Write-Host "Copy tool           : $copyTool"
Write-Host "Preserve structure  : $preserveTree"
Write-Host "AutoRun             : $autoRun"
Write-Host "------------------------------------------------------------"

# ============================================================
# FIND CSV FILES
# ============================================================

$csvFiles = @(
    Get-ChildItem `
        -LiteralPath $reportFolder `
        -Filter "*.csv" `
        -File |
    Sort-Object Name
)

if ($csvFiles.Count -eq 0) {
    throw "No CSV files found in: $reportFolder"
}

Write-Host ""
Write-Host "CSV files found: $($csvFiles.Count)"

foreach ($csv in $csvFiles) {
    Write-Host "  $($csv.Name)"
}

# ============================================================
# UNIQUE FILE HASHSET
# ============================================================

$uniqueFiles = New-Object `
    System.Collections.Generic.HashSet[string] `
    ([System.StringComparer]::OrdinalIgnoreCase)

# ============================================================
# STATISTICS
# ============================================================

$totalRows = 0
$processedFiles = 0
$skippedFiles = 0

# ============================================================
# PROCESS CSV FILES
# ============================================================

foreach ($csvFile in $csvFiles) {

    Write-Host ""
    Write-Host "============================================================"
    Write-Host "Processing: $($csvFile.Name)"
    Write-Host "============================================================"

    $reader = $null

    try {

        # ----------------------------------------------------
        # Open file as stream
        # ----------------------------------------------------

        $reader = New-Object System.IO.StreamReader(
            $csvFile.FullName,
            [System.Text.Encoding]::Default
        )

        # ----------------------------------------------------
        # Read header
        # ----------------------------------------------------

        $header = $reader.ReadLine()

        if ([string]::IsNullOrWhiteSpace($header)) {

            Write-Host "CSV is empty. Skipping." `
                -ForegroundColor Yellow

            $skippedFiles++
            continue
        }

        # ----------------------------------------------------
        # Detect delimiter
        # ----------------------------------------------------

        $delimiter = Find-Delimiter `
            -Header $header `
            -ParentColumn $parentCol `
            -ChildColumn $childCol

        if ($null -eq $delimiter) {

            Write-Host ""
            Write-Host "Could not determine delimiter." `
                -ForegroundColor Yellow

            Write-Host "Skipping: $($csvFile.Name)"

            $skippedFiles++
            continue
        }

        if ($delimiter -eq "`t") {
            Write-Host "Delimiter detected : TAB"
        }
        else {
            Write-Host "Delimiter detected : '$delimiter'"
        }

        # ----------------------------------------------------
        # Parse header
        # ----------------------------------------------------

        $headers = Parse-DelimitedLine `
            -Line $header `
            -Delimiter $delimiter

        # ----------------------------------------------------
        # Locate columns dynamically
        # ----------------------------------------------------

        $parentIndex = -1
        $childIndex = -1

        for (
            $i = 0;
            $i -lt $headers.Count;
            $i++
        ) {

            $columnName = Remove-OuterQuotes $headers[$i]

            if (
                $columnName.Equals(
                    $parentCol,
                    [System.StringComparison]::OrdinalIgnoreCase
                )
            ) {
                $parentIndex = $i
            }

            if (
                $columnName.Equals(
                    $childCol,
                    [System.StringComparison]::OrdinalIgnoreCase
                )
            ) {
                $childIndex = $i
            }
        }

        if (
            $parentIndex -lt 0 -or
            $childIndex -lt 0
        ) {

            Write-Host ""
            Write-Host "Required columns not found." `
                -ForegroundColor Yellow

            Write-Host ""
            Write-Host "Columns found:"

            foreach ($column in $headers) {
                Write-Host "  $(Remove-OuterQuotes $column)"
            }

            $skippedFiles++
            continue
        }

        Write-Host "ParentPath column   : $parentIndex"
        Write-Host "ChildPath column    : $childIndex"

        # ----------------------------------------------------
        # Process each row
        # ----------------------------------------------------

        $fileRowCount = 0

        while (-not $reader.EndOfStream) {

            $line = $reader.ReadLine()

            if ([string]::IsNullOrWhiteSpace($line)) {
                continue
            }

            $values = Parse-DelimitedLine `
                -Line $line `
                -Delimiter $delimiter

            $fileRowCount++
            $totalRows++

            # ------------------------------------------------
            # ParentPath
            # ------------------------------------------------

            if ($parentIndex -lt $values.Count) {

                $parentPath = Normalize-FilePath `
                    $values[$parentIndex]

                if (-not [string]::IsNullOrWhiteSpace($parentPath)) {

                    [void]$uniqueFiles.Add($parentPath)
                }
            }

            # ------------------------------------------------
            # ChildPath
            # ------------------------------------------------

            if ($childIndex -lt $values.Count) {

                $childPath = Normalize-FilePath `
                    $values[$childIndex]

                if (-not [string]::IsNullOrWhiteSpace($childPath)) {

                    [void]$uniqueFiles.Add($childPath)
                }
            }
        }

        $processedFiles++

        Write-Host "Rows processed      : $fileRowCount"
        Write-Host "Unique files so far : $($uniqueFiles.Count)"
    }
    catch {

        Write-Host ""
        Write-Host "ERROR processing $($csvFile.Name)" `
            -ForegroundColor Red

        Write-Host $_.Exception.Message `
            -ForegroundColor Red

        $skippedFiles++
    }
    finally {

        # ----------------------------------------------------
        # Close current CSV immediately.
        # ----------------------------------------------------

        if ($null -ne $reader) {

            $reader.Close()
            $reader.Dispose()
            $reader = $null
        }

        # Explicitly release temporary variables
        $header = $null
        $headers = $null
        $values = $null
    }
}

# ============================================================
# SUMMARY
# ============================================================

Write-Host ""
Write-Host "============================================================"
Write-Host " BOM PROCESSING SUMMARY"
Write-Host "============================================================"
Write-Host "CSV files found     : $($csvFiles.Count)"
Write-Host "CSV files processed : $processedFiles"
Write-Host "CSV files skipped   : $skippedFiles"
Write-Host "Total BOM rows      : $totalRows"
Write-Host "Unique files        : $($uniqueFiles.Count)"
Write-Host "============================================================"

# ============================================================
# OUTPUT FILES
# ============================================================

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

$batchFile = Join-Path `
    $reportFolder `
    "CopyBOMFiles_$timestamp.bat"

$logFile = Join-Path `
    $reportFolder `
    "CopyBOMFiles_$timestamp.log"

$missingLog = Join-Path `
    $reportFolder `
    "MissingFiles_$timestamp.txt"

# ============================================================
# CREATE BATCH FILE
# ============================================================

$batchLines = New-Object `
    System.Collections.Generic.List[string]

$batchLines.Add("@echo off")
$batchLines.Add("REM ============================================================")
$batchLines.Add("REM SolidWorks BOM Copy Batch")
$batchLines.Add("REM Generated by Generate-CopyBatch.ps1")
$batchLines.Add("REM Generated: $(Get-Date)")
$batchLines.Add("REM Copy Tool: $copyTool")
$batchLines.Add("REM ============================================================")
$batchLines.Add("")
$batchLines.Add(
    "echo Copy run started %date% %time% > `"$logFile`""
)
$batchLines.Add("")

# ============================================================
# DIRECTORY TRACKING
# ============================================================

$createdDirectories = New-Object `
    System.Collections.Generic.HashSet[string] `
    ([System.StringComparer]::OrdinalIgnoreCase)

# ============================================================
# MISSING FILES
# ============================================================

$missing = New-Object `
    System.Collections.Generic.List[string]

$copyCount = 0

# ============================================================
# GENERATE COPY COMMANDS
# ============================================================

foreach ($file in $uniqueFiles) {

    # --------------------------------------------------------
    # Check source file
    # --------------------------------------------------------

    if (
        -not (Test-Path `
            -LiteralPath $file `
            -PathType Leaf)
    ) {

        $missing.Add($file)

        continue
    }

    $copyCount++

    # --------------------------------------------------------
    # Source directory
    # --------------------------------------------------------

    $sourceDir = Split-Path `
        -Path $file `
        -Parent

    # --------------------------------------------------------
    # Destination directory
    # --------------------------------------------------------

    if ($preserveTree) {

        $relativeDir =
            $sourceDir -replace '^[A-Za-z]:', ''

        $relativeDir =
            $relativeDir.TrimStart('\')

        $targetDir = Join-Path `
            $destPath `
            $relativeDir
    }
    else {

        $targetDir = $destPath
    }

    # --------------------------------------------------------
    # Create destination directory only once
    # --------------------------------------------------------

    if ($createdDirectories.Add($targetDir)) {

        $batchLines.Add(
            "if not exist `"$targetDir`" mkdir `"$targetDir`""
        )
    }

    # --------------------------------------------------------
    # XCOPY
    # --------------------------------------------------------

    if ($copyTool -eq "XCOPY") {

        $batchLines.Add(
            "xcopy `"$file`" `"$targetDir\`" $xcopySw >> `"$logFile`" 2>&1"
        )
    }

    # --------------------------------------------------------
    # ROBOCOPY
    # --------------------------------------------------------

    else {

        $fileName = Split-Path `
            -Path $file `
            -Leaf

        $batchLines.Add(
            "robocopy `"$sourceDir`" `"$targetDir`" `"$fileName`" $robocopySw >> `"$logFile`" 2>&1"
        )
    }
}

# ============================================================
# COMPLETE BATCH
#
# IMPORTANT:
# No "pause" command.
# Therefore the BAT closes automatically when complete.
# ============================================================

$batchLines.Add("")
$batchLines.Add(
    "echo Copy run finished %date% %time% >> `"$logFile`""
)

$batchLines.Add(
    "echo Files attempted: $copyCount >> `"$logFile`""
)

$batchLines.Add(
    "echo Missing files: $($missing.Count) >> `"$logFile`""
)

# ============================================================
# WRITE BAT
# ============================================================

$batchLines |
    Set-Content `
        -LiteralPath $batchFile `
        -Encoding ASCII

# ============================================================
# WRITE MISSING FILE REPORT
# ============================================================

if ($missing.Count -gt 0) {

    $missing |
        Sort-Object |
        Out-File `
            -LiteralPath $missingLog `
            -Encoding ASCII
}

# ============================================================
# FINAL OUTPUT
# ============================================================

Write-Host ""
Write-Host "============================================================"
Write-Host " GENERATION COMPLETE"
Write-Host "============================================================"
Write-Host ""
Write-Host "Batch file:"
Write-Host "$batchFile"
Write-Host ""
Write-Host "Files to copy : $copyCount"
Write-Host "Missing files : $($missing.Count)"
Write-Host ""

if ($missing.Count -gt 0) {

    Write-Host "Missing files report:"
    Write-Host "$missingLog" `
        -ForegroundColor Yellow

    Write-Host ""
}

# ============================================================
# AUTORUN
# ============================================================

if ($autoRun) {

    Write-Host "AutoRun=true"
    Write-Host "Executing generated batch file..."
    Write-Host ""

    Start-Process `
        -FilePath $batchFile `
        -Wait `
        -NoNewWindow

    Write-Host ""
    Write-Host "Copy operation completed."
    Write-Host ""
    Write-Host "Log:"
    Write-Host "$logFile"
}
else {

    Write-Host "AutoRun=false"
    Write-Host ""
    Write-Host "The batch file was generated but not executed."
    Write-Host ""
    Write-Host "Review the BAT file and execute it manually."
}

Write-Host ""
Write-Host "============================================================"