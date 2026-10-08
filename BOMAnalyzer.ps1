#requires -Version 5.1
# GIT Check
<#
============================================================
 BOM CONSOLIDATION AND ANALYSIS UTILITY
============================================================

Reports generated:

    Consolidated_BOM.csv
    DuplicateOccurrences.csv
    MissingReferences.csv
    VirtualBOMLines.csv
    InvalidPaths.csv
    RunSummary.txt


Duplicate Occurrences
---------------------
Same ChildName exists with different ChildPath values.

The report shows:

    ChildName
    ChildPath
    AlternatePaths
    OccurrenceCount
    SourceFile

AlternatePaths contains ONLY the paths other than the
ChildPath shown on that particular row.


Missing References
------------------
ParentPath and/or ChildPath does not physically exist.

Mapped drives are checked before Test-Path is performed.


Virtual BOM Lines
-----------------
IsVirtual is:

    true
    1
    yes
    y


Invalid Paths
-------------
A path is invalid if it does not belong to:

    S:\R&D-Partage\PRODUITS
    S:\R&D-Partage\FOURNISSEURS
    S:\R&D-Partage\PIECES

or:

    Z:\PRODUITS
    Z:\FOURNISSEURS
    Z:\PIECES


Performance
-----------
File existence is cached by unique path.
============================================================
#>

Set-StrictMode -Version Latest

$ErrorActionPreference = "Stop"


# ============================================================
# SCRIPT ROOT
# ============================================================

$ScriptRoot =
    Split-Path -Parent $MyInvocation.MyCommand.Definition

$ConfigPath =
    Join-Path $ScriptRoot "config.ini"


# ============================================================
# READ INI
# ============================================================

function Get-IniFile {

    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $Ini = @{}

    $Section = ""

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Configuration file not found: $Path"
    }

    $Lines = @(
        Get-Content -LiteralPath $Path
    )

    foreach ($LineObject in @($Lines)) {

        $Line =
            ([string]$LineObject).Trim()

        if ([string]::IsNullOrWhiteSpace($Line)) {
            continue
        }

        if (
            $Line.StartsWith(";") -or
            $Line.StartsWith("#")
        ) {
            continue
        }

        if ($Line -match '^\[(.+)\]$') {

            $Section =
                $matches[1].Trim()

            if (-not $Ini.ContainsKey($Section)) {
                $Ini[$Section] = @{}
            }

            continue
        }

        if ($Line -match '^([^=]+)=(.*)$') {

            $Key =
                $matches[1].Trim()

            $Value =
                $matches[2].Trim()

            if (-not $Ini.ContainsKey($Section)) {
                $Ini[$Section] = @{}
            }

            $Ini[$Section][$Key] =
                $Value
        }
    }

    return $Ini
}


# ============================================================
# LOAD CONFIGURATION
# ============================================================

$Config =
    Get-IniFile -Path $ConfigPath


$InputFolder =
    Join-Path `
        $ScriptRoot `
        $Config["Paths"]["InputFolder"]


$OutputFolder =
    Join-Path `
        $ScriptRoot `
        $Config["Paths"]["OutputFolder"]


$SharedDriveRoot =
    $Config["Settings"]["SharedDriveRoot"]


$AlternateRoots = @(
    $Config["Settings"]["AlternateRoots"] -split "," |
    ForEach-Object {
        ([string]$_).Trim().TrimEnd("\")
    } |
    Where-Object {
        -not [string]::IsNullOrWhiteSpace($_)
    }
)


$AllowedRootFolders = @(
    $Config["Settings"]["AllowedRootFolders"] -split "," |
    ForEach-Object {
        ([string]$_).Trim().Trim("\")
    } |
    Where-Object {
        -not [string]::IsNullOrWhiteSpace($_)
    }
)


# ============================================================
# BUILD VALID ROOT PATHS
# ============================================================

$ValidRootPaths =
    [System.Collections.Generic.List[string]]::new()


$SharedRoot =
    $SharedDriveRoot.TrimEnd("\")


foreach ($Folder in @($AllowedRootFolders)) {

    $ValidRootPaths.Add(
        "$SharedRoot\$Folder"
    )
}


foreach ($AlternateRoot in @($AlternateRoots)) {

    foreach ($Folder in @($AllowedRootFolders)) {

        $Root =
            $AlternateRoot.TrimEnd("\")

        $ValidRootPaths.Add(
            "$Root\$Folder"
        )
    }
}


$ValidRootPaths = @(
    $ValidRootPaths |
    Sort-Object -Unique
)


# ============================================================
# CREATE OUTPUT FOLDER
# ============================================================

if (-not (Test-Path -LiteralPath $OutputFolder)) {

    New-Item `
        -ItemType Directory `
        -Path $OutputFolder `
        -Force |
        Out-Null
}


# ============================================================
# COLUMN ALIASES
# ============================================================

function Get-ColumnAliases {

    param(
        [AllowNull()]
        [string]$ConfigValue
    )

    if ([string]::IsNullOrWhiteSpace($ConfigValue)) {
        return @()
    }

    return @(
        $ConfigValue -split "," |
        ForEach-Object {
            ([string]$_).Trim()
        } |
        Where-Object {
            -not [string]::IsNullOrWhiteSpace($_)
        }
    )
}


$ParentPathAliases =
    Get-ColumnAliases `
        $Config["Columns"]["ParentPath"]


$ChildPathAliases =
    Get-ColumnAliases `
        $Config["Columns"]["ChildPath"]


$ChildNameAliases =
    Get-ColumnAliases `
        $Config["Columns"]["ChildName"]


$IsVirtualAliases =
    Get-ColumnAliases `
        $Config["Columns"]["IsVirtual"]


# ============================================================
# NORMALIZE COLUMN NAME
# ============================================================

function Normalize-ColumnName {

    param(
        [AllowNull()]
        [string]$Name
    )

    if ($null -eq $Name) {
        return ""
    }

    $Result =
        ([string]$Name).Trim()

    $Result =
        $Result -replace '[\s_\-\.]', ""

    $Result =
        $Result -replace '[^a-zA-Z0-9]', ""

    return $Result.ToLowerInvariant()
}


# ============================================================
# FIND COLUMN
# ============================================================

function Find-ColumnName {

    param(
        [AllowNull()]
        [string[]]$Headers,

        [AllowNull()]
        [string[]]$Aliases
    )

    if (
        $null -eq $Headers -or
        $null -eq $Aliases
    ) {
        return $null
    }

    $NormalizedAliases = @(
        @($Aliases) |
        ForEach-Object {
            Normalize-ColumnName $_
        }
    )

    foreach ($Header in @($Headers)) {

        $NormalizedHeader =
            Normalize-ColumnName $Header

        if (
            $NormalizedAliases -contains
            $NormalizedHeader
        ) {
            return $Header
        }
    }

    return $null
}


# ============================================================
# SAFE ROW VALUE
# ============================================================

function Get-RowValue {

    param(
        [AllowNull()]
        [object]$Row,

        [AllowNull()]
        [string]$ColumnName
    )

    if (
        $null -eq $Row -or
        [string]::IsNullOrWhiteSpace($ColumnName)
    ) {
        return ""
    }

    $Property =
        $Row.PSObject.Properties[$ColumnName]

    if (
        $null -eq $Property -or
        $null -eq $Property.Value
    ) {
        return ""
    }

    return ([string]$Property.Value)
}


# ============================================================
# DELIMITER CANDIDATES
# ============================================================

function Get-DelimiterCandidates {

    return @(
        "##"
        "||"
        "@@"
        "::"
        ";;"
        "`t"
        "|"
        ","
        ";"
    )
}


# ============================================================
# FIND CSV DELIMITER
# ============================================================

function Find-Delimiter {

    param(
        [AllowNull()]
        [string]$HeaderLine
    )

    if ([string]::IsNullOrWhiteSpace($HeaderLine)) {
        return ","
    }

    $BestDelimiter = $null
    $BestCount = 0

    foreach ($Delimiter in @(Get-DelimiterCandidates)) {

        $Count = 0
        $InsideQuotes = $false
        $Index = 0

        while ($Index -lt $HeaderLine.Length) {

            $Character =
                $HeaderLine[$Index]

            if ($Character -eq '"') {

                if (
                    $InsideQuotes -and
                    ($Index + 1 -lt $HeaderLine.Length) -and
                    ($HeaderLine[$Index + 1] -eq '"')
                ) {

                    $Index += 2
                    continue
                }

                $InsideQuotes =
                    -not $InsideQuotes

                $Index++
                continue
            }

            if (
                -not $InsideQuotes -and
                ($Index + $Delimiter.Length) -le
                $HeaderLine.Length
            ) {

                $PossibleDelimiter =
                    $HeaderLine.Substring(
                        $Index,
                        $Delimiter.Length
                    )

                if ($PossibleDelimiter -eq $Delimiter) {

                    $Count++

                    $Index +=
                        $Delimiter.Length

                    continue
                }
            }

            $Index++
        }

        if ($Count -gt $BestCount) {

            $BestCount =
                $Count

            $BestDelimiter =
                $Delimiter
        }
    }

    if ($null -eq $BestDelimiter) {
        return ","
    }

    return $BestDelimiter
}


# ============================================================
# PARSE CSV LINE
# ============================================================

function Parse-CsvLine {

    param(
        [AllowNull()]
        [string]$Line,

        [AllowNull()]
        [string]$Delimiter
    )

    if ($null -eq $Line) {
        return @("")
    }

    if ([string]::IsNullOrEmpty($Delimiter)) {
        $Delimiter = ","
    }

    $Fields =
        [System.Collections.Generic.List[string]]::new()

    $Current =
        New-Object System.Text.StringBuilder

    $InsideQuotes = $false
    $Index = 0

    while ($Index -lt $Line.Length) {

        $Character =
            $Line[$Index]

        if ($Character -eq '"') {

            if (
                $InsideQuotes -and
                ($Index + 1 -lt $Line.Length) -and
                ($Line[$Index + 1] -eq '"')
            ) {

                [void]$Current.Append('"')

                $Index += 2

                continue
            }

            $InsideQuotes =
                -not $InsideQuotes

            $Index++

            continue
        }

        if (
            -not $InsideQuotes -and
            ($Index + $Delimiter.Length) -le
            $Line.Length
        ) {

            $PossibleDelimiter =
                $Line.Substring(
                    $Index,
                    $Delimiter.Length
                )

            if ($PossibleDelimiter -eq $Delimiter) {

                $Fields.Add(
                    $Current.ToString()
                )

                [void]$Current.Clear()

                $Index +=
                    $Delimiter.Length

                continue
            }
        }

        [void]$Current.Append(
            $Character
        )

        $Index++
    }

    $Fields.Add(
        $Current.ToString()
    )

    return @(
        $Fields.ToArray()
    )
}


# ============================================================
# READ DYNAMIC CSV
# ============================================================

function Read-DynamicCsv {

    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )

    $Lines = @(
        Get-Content -LiteralPath $FilePath
    )

    if ($Lines.Count -eq 0) {

        return [PSCustomObject]@{
            Headers   = @()
            Records   = @()
            Delimiter = ","
            File      = $FilePath
        }
    }

    $Delimiter =
        Find-Delimiter `
            -HeaderLine ([string]$Lines[0])


    $Headers = @(
        Parse-CsvLine `
            -Line ([string]$Lines[0]) `
            -Delimiter $Delimiter |
        ForEach-Object {
            ([string]$_).Trim()
        }
    )


    for (
        $HeaderIndex = 0;
        $HeaderIndex -lt $Headers.Count;
        $HeaderIndex++
    ) {

        if (
            [string]::IsNullOrWhiteSpace(
                $Headers[$HeaderIndex]
            )
        ) {

            $Headers[$HeaderIndex] =
                "Column_$($HeaderIndex + 1)"
        }
    }


    $Records =
        [System.Collections.Generic.List[object]]::new()


    for (
        $LineIndex = 1;
        $LineIndex -lt $Lines.Count;
        $LineIndex++
    ) {

        $Line =
            [string]$Lines[$LineIndex]


        if (
            [string]::IsNullOrWhiteSpace($Line)
        ) {
            continue
        }


        $Values = @(
            Parse-CsvLine `
                -Line $Line `
                -Delimiter $Delimiter
        )


        $Record =
            [ordered]@{}


        for (
            $ColumnIndex = 0;
            $ColumnIndex -lt $Headers.Count;
            $ColumnIndex++
        ) {

            $Value = ""


            if (
                $ColumnIndex -lt $Values.Count
            ) {

                $Value =
                    [string]$Values[$ColumnIndex]
            }


            $Record[
                $Headers[$ColumnIndex]
            ] =
                $Value
        }


        $Record["_SourceFile"] =
            [System.IO.Path]::GetFileName(
                $FilePath
            )


        $Records.Add(
            [PSCustomObject]$Record
        )
    }


    return [PSCustomObject]@{
        Headers   = @($Headers)
        Records   = @($Records)
        Delimiter = $Delimiter
        File      = $FilePath
    }
}


# ============================================================
# PATH NORMALIZATION CACHE
# ============================================================

$PathNormalizationCache =
    [System.Collections.Generic.Dictionary[string,string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )


function Get-NormalizedPathCached {

    param(
        [AllowNull()]
        [string]$Path
    )

    if ($null -eq $Path) {
        return ""
    }

    if (
        $PathNormalizationCache.ContainsKey($Path)
    ) {

        return $PathNormalizationCache[$Path]
    }


    $Result =
        $Path.Trim()


    if (
        [string]::IsNullOrWhiteSpace($Result)
    ) {

        $PathNormalizationCache[$Path] = ""

        return ""
    }


    $Result =
        $Result.Trim('"').Trim()


    $Result =
        $Result.Replace("/", "\")


    $Result =
        $Result.TrimEnd("\")


    $PathNormalizationCache[$Path] =
        $Result


    return $Result
}


# ============================================================
# VALID PATH CACHE
# ============================================================

$ValidPathCache =
    [System.Collections.Generic.Dictionary[string,bool]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )


function Test-IsAllowedBOMPathCached {

    param(
        [AllowNull()]
        [string]$Path
    )

    if (
        [string]::IsNullOrWhiteSpace($Path)
    ) {
        return $false
    }


    if (
        $ValidPathCache.ContainsKey($Path)
    ) {

        return $ValidPathCache[$Path]
    }


    $IsValid =
        $false


    foreach ($ValidRoot in @($ValidRootPaths)) {

        $Root =
            ([string]$ValidRoot).TrimEnd("\")


        if (
            $Path.Equals(
                $Root,
                [StringComparison]::OrdinalIgnoreCase
            )
        ) {

            $IsValid =
                $true

            break
        }


        if (
            $Path.StartsWith(
                "$Root\",
                [StringComparison]::OrdinalIgnoreCase
            )
        ) {

            $IsValid =
                $true

            break
        }
    }


    $ValidPathCache[$Path] =
        $IsValid


    return $IsValid
}


# ============================================================
# DRIVE AVAILABILITY CACHE
# ============================================================

$DriveAvailabilityCache =
    [System.Collections.Generic.Dictionary[string,bool]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )


function Get-DriveLetterFromPath {

    param(
        [AllowNull()]
        [string]$Path
    )

    if (
        [string]::IsNullOrWhiteSpace($Path)
    ) {
        return $null
    }


    if (
        $Path -match '^(?<Drive>[A-Za-z]):\\'
    ) {

        return (
            $matches["Drive"].ToUpperInvariant()
        )
    }


    return $null
}


function Test-DriveAvailable {

    param(
        [AllowNull()]
        [string]$DriveLetter
    )

    if (
        [string]::IsNullOrWhiteSpace($DriveLetter)
    ) {
        return $false
    }


    $DriveLetter =
        $DriveLetter.Trim().TrimEnd(":")


    if (
        $DriveAvailabilityCache.ContainsKey(
            $DriveLetter
        )
    ) {

        return $DriveAvailabilityCache[
            $DriveLetter
        ]
    }


    $DriveRoot =
        "$DriveLetter`:\"


    $Available =
        $false


    try {

        $Available =
            Test-Path `
                -LiteralPath $DriveRoot `
                -ErrorAction SilentlyContinue
    }
    catch {

        $Available =
            $false
    }


    if (-not $Available) {

        try {

            $PSDrive =
                Get-PSDrive `
                    -Name $DriveLetter `
                    -ErrorAction SilentlyContinue


            if ($null -ne $PSDrive) {

                try {

                    $Available =
                        Test-Path `
                            -LiteralPath $PSDrive.Root `
                            -ErrorAction SilentlyContinue
                }
                catch {

                    $Available =
                        $false
                }
            }
        }
        catch {

            $Available =
                $false
        }
    }


    $DriveAvailabilityCache[
        $DriveLetter
    ] =
        [bool]$Available


    return [bool]$Available
}


# ============================================================
# FILE EXISTENCE CACHE
# ============================================================

$FileExistsCache =
    [System.Collections.Generic.Dictionary[string,bool]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )


function Test-FileExistsCached {

    param(
        [AllowNull()]
        [string]$Path
    )

    if (
        [string]::IsNullOrWhiteSpace($Path)
    ) {
        return $false
    }


    if (
        $FileExistsCache.ContainsKey($Path)
    ) {

        return $FileExistsCache[$Path]
    }


    $DriveLetter =
        Get-DriveLetterFromPath `
            -Path $Path


    if (
        $null -ne $DriveLetter
    ) {

        $DriveAvailable =
            Test-DriveAvailable `
                -DriveLetter $DriveLetter


        if (-not $DriveAvailable) {

            return $null
        }
    }


    $Exists =
        $false


    try {

        $Exists =
            Test-Path `
                -LiteralPath $Path `
                -ErrorAction SilentlyContinue
    }
    catch {

        $Exists =
            $false
    }


    $FileExistsCache[$Path] =
        [bool]$Exists


    return [bool]$Exists
}


# ============================================================
# GET CSV FILES
# ============================================================

$CsvFiles = @(
    Get-ChildItem `
        -LiteralPath $InputFolder `
        -Filter "*.csv" `
        -File
)


if (
    $CsvFiles.Count -eq 0
) {

    throw (
        "No CSV files found in: " +
        $InputFolder
    )
}


# ============================================================
# START
# ============================================================

Write-Host ""

Write-Host `
    "============================================" `
    -ForegroundColor Cyan

Write-Host `
    " BOM CONSOLIDATION AND ANALYSIS UTILITY" `
    -ForegroundColor Cyan

Write-Host `
    "============================================" `
    -ForegroundColor Cyan

Write-Host ""

Write-Host "Input Folder : $InputFolder"
Write-Host "Output Folder: $OutputFolder"
Write-Host "CSV Files    : $($CsvFiles.Count)"

Write-Host ""


# ============================================================
# CONSOLIDATE CSV FILES
# ============================================================

$AllHeaders =
    [System.Collections.Generic.List[string]]::new()


$AllRows =
    [System.Collections.Generic.List[object]]::new()


foreach ($CsvFile in @($CsvFiles)) {

    Write-Host `
        "Reading: $($CsvFile.Name)" `
        -ForegroundColor Yellow


    $CsvData =
        Read-DynamicCsv `
            -FilePath $CsvFile.FullName


    foreach ($Header in @($CsvData.Headers)) {

        if (
            -not (
                $AllHeaders -contains
                $Header
            )
        ) {

            $AllHeaders.Add(
                $Header
            )
        }
    }


    # ========================================================
    # DETECT TOP-LINE FOR THIS CSV
    #
    # The first BOM record in each CSV is the assembly
    # top-line. Its ParentPath is expected to be ROOT.
    # The ChildName of that first record is copied into the
    # TopLine column for every record originating from this CSV.
    # ========================================================

    $CsvParentPathColumn =
        Find-ColumnName `
            -Headers @($CsvData.Headers) `
            -Aliases @($ParentPathAliases)

    $CsvChildNameColumn =
        Find-ColumnName `
            -Headers @($CsvData.Headers) `
            -Aliases @($ChildNameAliases)

    $TopLine = ""

    $CsvRecords = @($CsvData.Records)

    if ($CsvRecords.Count -gt 0) {

        $FirstRow = $CsvRecords[0]
        $FirstParentPath = ""

        if ($null -ne $CsvParentPathColumn) {
            $FirstParentPath =
                Get-RowValue `
                    -Row $FirstRow `
                    -ColumnName $CsvParentPathColumn
        }

        # The first record is the top-line when its ParentPath
        # is ROOT. Keep the ChildName exactly as supplied by
        # the source CSV.
        if (
            $FirstParentPath.Trim().Equals(
                "ROOT",
                [StringComparison]::OrdinalIgnoreCase
            )
        ) {

            if ($null -ne $CsvChildNameColumn) {
                $TopLine =
                    Get-RowValue `
                        -Row $FirstRow `
                        -ColumnName $CsvChildNameColumn
            }
        }
        else {
            Write-Host `
                "WARNING: First row of $($CsvFile.Name) does not have ParentPath = ROOT. TopLine will be blank for this CSV." `
                -ForegroundColor Yellow
        }
    }
    else {
        Write-Host `
            "WARNING: $($CsvFile.Name) contains no BOM records. TopLine cannot be determined." `
            -ForegroundColor Yellow
    }


    foreach ($SourceRow in @($CsvRecords)) {

        # Add/overwrite the generated TopLine value on every
        # row belonging to this source CSV.
        $null =
            $SourceRow |
            Add-Member `
                -MemberType NoteProperty `
                -Name "TopLine" `
                -Value $TopLine `
                -Force

        $AllRows.Add(
            $SourceRow
        )
    }
}


if (
    -not (
        $AllHeaders -contains
        "TopLine"
    )
) {

    $AllHeaders.Add(
        "TopLine"
    )
}


if (
    -not (
        $AllHeaders -contains
        "_SourceFile"
    )
) {

    $AllHeaders.Add(
        "_SourceFile"
    )
}


# ============================================================
# WRITE CONSOLIDATED BOM
# ============================================================

$ConsolidatedPath =
    Join-Path `
        $OutputFolder `
        "Consolidated_BOM.csv"


$ConsolidatedRows =
    [System.Collections.Generic.List[object]]::new()


foreach ($Row in @($AllRows)) {

    $OutputRow =
        [ordered]@{}


    foreach ($Header in @($AllHeaders)) {

        $OutputRow[$Header] =
            Get-RowValue `
                -Row $Row `
                -ColumnName $Header
    }


    $ConsolidatedRows.Add(
        [PSCustomObject]$OutputRow
    )
}


if (
    $ConsolidatedRows.Count -gt 0
) {

    $ConsolidatedRows |
        Export-Csv `
            -LiteralPath $ConsolidatedPath `
            -NoTypeInformation `
            -Encoding UTF8
}
else {

    Set-Content `
        -LiteralPath $ConsolidatedPath `
        -Value "" `
        -Encoding UTF8
}


Write-Host ""

Write-Host `
    "Created: Consolidated_BOM.csv" `
    -ForegroundColor Green


# ============================================================
# DETECT COLUMNS
# ============================================================

$ParentPathColumn =
    Find-ColumnName `
        -Headers @($AllHeaders) `
        -Aliases @($ParentPathAliases)


$ChildPathColumn =
    Find-ColumnName `
        -Headers @($AllHeaders) `
        -Aliases @($ChildPathAliases)


$ChildNameColumn =
    Find-ColumnName `
        -Headers @($AllHeaders) `
        -Aliases @($ChildNameAliases)


$IsVirtualColumn =
    Find-ColumnName `
        -Headers @($AllHeaders) `
        -Aliases @($IsVirtualAliases)


# ============================================================
# CREATE INTERNAL ANALYSIS ROWS
# ============================================================

$AnalysisRows =
    [System.Collections.Generic.List[object]]::new()


foreach ($Row in @($AllRows)) {

    $ParentPathRaw = ""
    $ChildPathRaw = ""
    $ChildName = ""
    $IsVirtualValue = ""


    if ($null -ne $ParentPathColumn) {

        $ParentPathRaw =
            Get-RowValue `
                -Row $Row `
                -ColumnName $ParentPathColumn
    }


    if ($null -ne $ChildPathColumn) {

        $ChildPathRaw =
            Get-RowValue `
                -Row $Row `
                -ColumnName $ChildPathColumn
    }


    if ($null -ne $ChildNameColumn) {

        $ChildName =
            (
                Get-RowValue `
                    -Row $Row `
                    -ColumnName $ChildNameColumn
            ).Trim()
    }


    if ($null -ne $IsVirtualColumn) {

        $IsVirtualValue =
            (
                Get-RowValue `
                    -Row $Row `
                    -ColumnName $IsVirtualColumn
            ).Trim().ToLowerInvariant()
    }


    $ParentPath =
        Get-NormalizedPathCached `
            -Path $ParentPathRaw


    $ChildPath =
        Get-NormalizedPathCached `
            -Path $ChildPathRaw


    if (
        [string]::IsNullOrWhiteSpace($ChildName) -and
        -not [string]::IsNullOrWhiteSpace($ChildPath)
    ) {

        try {

            $ChildName =
                [System.IO.Path]::GetFileName(
                    $ChildPath
                )
        }
        catch {

            $ChildName =
                $ChildPath
        }
    }


    $AnalysisRows.Add(

        [PSCustomObject]@{

            ParentPath =
                $ParentPath

            ChildPath =
                $ChildPath

            ChildName =
                $ChildName

            IsVirtualValue =
                $IsVirtualValue

            SourceFile =
                Get-RowValue `
                    -Row $Row `
                    -ColumnName "_SourceFile"
        }
    )
}


# ============================================================
# COLLECT UNIQUE PATHS
# ============================================================

$UniquePaths =
    [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )


foreach ($Row in @($AnalysisRows)) {

    if (
        -not [string]::IsNullOrWhiteSpace(
            $Row.ParentPath
        )
    ) {

        [void]$UniquePaths.Add(
            $Row.ParentPath
        )
    }


    if (
        -not [string]::IsNullOrWhiteSpace(
            $Row.ChildPath
        )
    ) {

        [void]$UniquePaths.Add(
            $Row.ChildPath
        )
    }
}


# ============================================================
# PRE-CHECK MAPPED DRIVES
# ============================================================

$DrivesUsedByBOM =
    [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )


foreach ($Path in @($UniquePaths)) {

    $DriveLetter =
        Get-DriveLetterFromPath `
            -Path $Path


    if ($null -ne $DriveLetter) {

        [void]$DrivesUsedByBOM.Add(
            $DriveLetter
        )
    }
}


Write-Host ""

Write-Host `
    "Checking drive availability..." `
    -ForegroundColor Cyan


foreach ($DriveLetter in @($DrivesUsedByBOM)) {

    $Available =
        Test-DriveAvailable `
            -DriveLetter $DriveLetter


    if ($Available) {

        Write-Host `
            "  $DriveLetter`: available" `
            -ForegroundColor Green
    }
    else {

        Write-Host `
            "  $DriveLetter`: unavailable - missing reference checks will be skipped for this drive" `
            -ForegroundColor Yellow
    }
}


# ============================================================
# CHECK UNIQUE FILE PATHS
# ============================================================

Write-Host ""

Write-Host `
    "Checking unique file references: $($UniquePaths.Count)" `
    -ForegroundColor Cyan


$PathCheckCounter =
    0


foreach ($Path in @($UniquePaths)) {

    $PathCheckCounter++


    $null =
        Test-FileExistsCached `
            -Path $Path


    if (
        ($PathCheckCounter % 500) -eq 0
    ) {

        Write-Host `
            "  Checked $PathCheckCounter / $($UniquePaths.Count)" `
            -ForegroundColor DarkGray
    }
}


Write-Host `
    "File existence checks completed." `
    -ForegroundColor Green


# ============================================================
# REPORT COLLECTIONS
# ============================================================

$DuplicateRows =
    [System.Collections.Generic.List[object]]::new()


$MissingRows =
    [System.Collections.Generic.List[object]]::new()


$VirtualRows =
    [System.Collections.Generic.List[object]]::new()


$InvalidPathRows =
    [System.Collections.Generic.List[object]]::new()


# ============================================================
# DUPLICATE GROUPS
# ============================================================

$ChildNameGroups =
    [System.Collections.Generic.Dictionary[
        string,
        System.Collections.Generic.List[object]
    ]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )


foreach ($Row in @($AnalysisRows)) {

    if (
        [string]::IsNullOrWhiteSpace(
            $Row.ChildName
        )
    ) {

        continue
    }


    $GroupKey =
        $Row.ChildName


    if (
        -not $ChildNameGroups.ContainsKey(
            $GroupKey
        )
    ) {

        $ChildNameGroups[$GroupKey] =
            [System.Collections.Generic.List[object]]::new()
    }


    $ChildNameGroups[$GroupKey].Add(
        $Row
    )
}


# ============================================================
# DUPLICATE OCCURRENCES
# ============================================================
#
# NEW BEHAVIOR:
#
# ChildPath:
#     Path used by this particular BOM occurrence.
#
# AlternatePaths:
#     ONLY the other paths used for this ChildName.
#
# ============================================================

foreach (
    $GroupKey in @($ChildNameGroups.Keys)
) {

    $Group =
        $ChildNameGroups[$GroupKey]


    $DifferentPaths =
        [System.Collections.Generic.HashSet[string]]::new(
            [System.StringComparer]::OrdinalIgnoreCase
        )


    foreach ($Row in @($Group)) {

        if (
            -not [string]::IsNullOrWhiteSpace(
                $Row.ChildPath
            )
        ) {

            [void]$DifferentPaths.Add(
                $Row.ChildPath
            )
        }
    }


    #
    # Duplicate means same ChildName but different paths.
    #

    if (
        $DifferentPaths.Count -gt 1
    ) {

        foreach ($Row in @($Group)) {

            #
            # Build alternate paths for THIS row.
            #

            $AlternatePaths =
                [System.Collections.Generic.List[string]]::new()


            foreach (
                $OtherPath in @($DifferentPaths)
            ) {

                if (
                    -not $OtherPath.Equals(
                        $Row.ChildPath,
                        [StringComparison]::OrdinalIgnoreCase
                    )
                ) {

                    $AlternatePaths.Add(
                        $OtherPath
                    )
                }
            }


            $DuplicateRows.Add(

                [PSCustomObject]@{

                    ChildName =
                        $Row.ChildName

                    ChildPath =
                        $Row.ChildPath

                    AlternatePaths =
                        (
                            @($AlternatePaths) -join
                            " | "
                        )

                    OccurrenceCount =
                        @($Group).Count

                    SourceFile =
                        $Row.SourceFile
                }
            )
        }
    }
}


# ============================================================
# MAIN ANALYSIS
# ============================================================

foreach ($Row in @($AnalysisRows)) {

    # ========================================================
    # MISSING REFERENCES
    # ========================================================

    $ParentExists =
        $null


    $ChildExists =
        $null


    if (
        -not [string]::IsNullOrWhiteSpace(
            $Row.ParentPath
        )
    ) {

        $ParentExists =
            Test-FileExistsCached `
                -Path $Row.ParentPath
    }


    if (
        -not [string]::IsNullOrWhiteSpace(
            $Row.ChildPath
        )
    ) {

        $ChildExists =
            Test-FileExistsCached `
                -Path $Row.ChildPath
    }


    #
    # Only confirmed FALSE values are treated as missing.
    #
    # NULL means drive unavailable and therefore cannot be
    # safely classified as missing.
    #

    $ParentConfirmedMissing =
        (
            ($null -ne $ParentExists) -and
            ($ParentExists -eq $false)
        )


    $ChildConfirmedMissing =
        (
            ($null -ne $ChildExists) -and
            ($ChildExists -eq $false)
        )


    if (
        $ParentConfirmedMissing -or
        $ChildConfirmedMissing
    ) {

        $MissingReference =
            [System.Collections.Generic.List[string]]::new()


        if ($ParentConfirmedMissing) {

            $MissingReference.Add(
                "Parent"
            )
        }


        if ($ChildConfirmedMissing) {

            $MissingReference.Add(
                "Child"
            )
        }


        $MissingRows.Add(

            [PSCustomObject]@{

                ParentPath =
                    $Row.ParentPath

                ChildPath =
                    $Row.ChildPath

                ParentPathExists =
                    $ParentExists

                ChildPathExists =
                    $ChildExists

                MissingReference =
                    (
                        $MissingReference -join
                        ", "
                    )

                SourceFile =
                    $Row.SourceFile
            }
        )
    }


    # ========================================================
    # VIRTUAL BOM LINE
    # ========================================================

    if (
        $Row.IsVirtualValue -in @(
            "true"
            "1"
            "yes"
            "y"
        )
    ) {

        $VirtualRows.Add(

            [PSCustomObject]@{

                ParentPath =
                    $Row.ParentPath

                ChildPath =
                    $Row.ChildPath

                IsVirtual =
                    $true

                SourceFile =
                    $Row.SourceFile
            }
        )
    }


    # ========================================================
    # INVALID PATHS
    # ========================================================

    if (
        -not [string]::IsNullOrWhiteSpace(
            $Row.ParentPath
        )
    ) {

        $ParentIsValid =
            Test-IsAllowedBOMPathCached `
                -Path $Row.ParentPath


        if (-not $ParentIsValid) {

            $InvalidPathRows.Add(

                [PSCustomObject]@{

                    PathType =
                        "Parent"

                    Path =
                        $Row.ParentPath

                    Reason =
                        "OutsideAllowedRootFolders"

                    SourceFile =
                        $Row.SourceFile
                }
            )
        }
    }


    if (
        -not [string]::IsNullOrWhiteSpace(
            $Row.ChildPath
        )
    ) {

        $ChildIsValid =
            Test-IsAllowedBOMPathCached `
                -Path $Row.ChildPath


        if (-not $ChildIsValid) {

            $InvalidPathRows.Add(

                [PSCustomObject]@{

                    PathType =
                        "Child"

                    Path =
                        $Row.ChildPath

                    Reason =
                        "OutsideAllowedRootFolders"

                    SourceFile =
                        $Row.SourceFile
                }
            )
        }
    }
}


# ============================================================
# WRITE REPORT FUNCTION
# ============================================================

function Write-Report {

    param(
        [AllowNull()]
        [object[]]$Rows,

        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Header
    )


    $Rows =
        @($Rows)


    if (
        $Rows.Count -gt 0
    ) {

        $Rows |
            Export-Csv `
                -LiteralPath $Path `
                -NoTypeInformation `
                -Encoding UTF8
    }
    else {

        Set-Content `
            -LiteralPath $Path `
            -Value $Header `
            -Encoding UTF8
    }
}


# ============================================================
# REPORT PATHS
# ============================================================

$DuplicatePath =
    Join-Path `
        $OutputFolder `
        "DuplicateOccurrences.csv"


$MissingPath =
    Join-Path `
        $OutputFolder `
        "MissingReferences.csv"


$VirtualPath =
    Join-Path `
        $OutputFolder `
        "VirtualBOMLines.csv"


$InvalidPath =
    Join-Path `
        $OutputFolder `
        "InvalidPaths.csv"


# ============================================================
# WRITE REPORTS
# ============================================================

Write-Report `
    -Rows @($DuplicateRows) `
    -Path $DuplicatePath `
    -Header "ChildName,ChildPath,AlternatePaths,OccurrenceCount,SourceFile"


Write-Report `
    -Rows @($MissingRows) `
    -Path $MissingPath `
    -Header "ParentPath,ChildPath,ParentPathExists,ChildPathExists,MissingReference,SourceFile"


Write-Report `
    -Rows @($VirtualRows) `
    -Path $VirtualPath `
    -Header "ParentPath,ChildPath,IsVirtual,SourceFile"


Write-Report `
    -Rows @($InvalidPathRows) `
    -Path $InvalidPath `
    -Header "PathType,Path,Reason,SourceFile"


# ============================================================
# COUNTS
# ============================================================

$InputCsvCount =
    @($CsvFiles).Count


$TotalBomRows =
    @($AnalysisRows).Count


$UniquePathCount =
    $UniquePaths.Count


$DuplicateCount =
    @($DuplicateRows).Count


$MissingCount =
    @($MissingRows).Count


$VirtualCount =
    @($VirtualRows).Count


$InvalidPathCount =
    @($InvalidPathRows).Count


$UnavailableDriveCount =
    @(
        $DrivesUsedByBOM |
        Where-Object {

            -not (
                Test-DriveAvailable `
                    -DriveLetter $_
            )
        }
    ).Count


# ============================================================
# RUN SUMMARY
# ============================================================

$SummaryPath =
    Join-Path `
        $OutputFolder `
        "RunSummary.txt"


$Summary = @"

BOM CONSOLIDATION UTILITY
=========================

Run Time:
$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")


INPUT
=====

Input CSV Files:
$InputCsvCount

Consolidated BOM Rows:
$TotalBomRows


PERFORMANCE
===========

Unique File Paths:
$UniquePathCount


REPORT COUNTS
=============

Duplicate Occurrences:
$DuplicateCount

Missing References:
$MissingCount

Virtual BOM Lines:
$VirtualCount

Invalid Paths:
$InvalidPathCount


MAPPED DRIVE STATUS
===================

Drives Used By BOM:
$($DrivesUsedByBOM -join ", ")

Unavailable Drives:
$UnavailableDriveCount


REPORT FILES
============

Consolidated_BOM.csv
DuplicateOccurrences.csv
MissingReferences.csv
VirtualBOMLines.csv
InvalidPaths.csv
RunSummary.txt

"@


Set-Content `
    -LiteralPath $SummaryPath `
    -Value $Summary `
    -Encoding UTF8


# ============================================================
# CLEANUP
# ============================================================

$AllRows = $null
$AnalysisRows = $null
$ChildNameGroups = $null
$DuplicateRows = $null
$MissingRows = $null
$VirtualRows = $null
$InvalidPathRows = $null


[GC]::Collect()
[GC]::WaitForPendingFinalizers()


# ============================================================
# FINAL OUTPUT
# ============================================================

Write-Host ""

Write-Host `
    "============================================" `
    -ForegroundColor Green

Write-Host `
    " PROCESSING COMPLETE" `
    -ForegroundColor Green

Write-Host `
    "============================================" `
    -ForegroundColor Green

Write-Host ""

Write-Host `
    "Output folder: $OutputFolder" `
    -ForegroundColor Cyan

Write-Host ""

Write-Host "Input CSV Files       : $InputCsvCount"
Write-Host "BOM Rows              : $TotalBomRows"
Write-Host "Unique Paths          : $UniquePathCount"
Write-Host "Duplicate Occurrences : $DuplicateCount"
Write-Host "Missing References    : $MissingCount"
Write-Host "Virtual BOM Lines     : $VirtualCount"
Write-Host "Invalid Paths         : $InvalidPathCount"

Write-Host ""

Write-Host "Created files:"

Write-Host "  Consolidated_BOM.csv"
Write-Host "  DuplicateOccurrences.csv"
Write-Host "  MissingReferences.csv"
Write-Host "  VirtualBOMLines.csv"
Write-Host "  InvalidPaths.csv"
Write-Host "  RunSummary.txt"

Write-Host ""

Write-Host `
    "Done." `
    -ForegroundColor Green

Write-Host ""