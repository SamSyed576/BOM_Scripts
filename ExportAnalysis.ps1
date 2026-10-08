# ============================================================
# Export-AnalysisReports.ps1
#
# Purpose:
#   Export SQL Server analysis report tables from StagingDB
#   into Excel-safe CSV files.
#
# Reports exported:
#   PC_101
#   PC_103
#   PC_104
#   PC_105
#   PC_106
#   PC_107
#   PC_108
#   PC_109
#   PC_110
#   PC_111
#
# CSV rules:
#   Column delimiter : ,
#   Text qualifier   : "
#   Multi-values     : |
#   Encoding         : UTF-8 with BOM
#
# The script explicitly quotes CSV fields so that:
#   - Pipe-separated values stay in one Excel cell
#   - Commas inside values do not create extra columns
#   - Double quotes inside values are escaped correctly
#   - Unicode characters are preserved
#
# ============================================================


# ============================================================
# CONFIGURATION
# ============================================================

$SqlServer = "TSRKTBLRL105\TCDB"
$Database  = "StagingDB"

# Output directory
$OutputFolder = "D:\Princecraft\Analysis reports"

# Windows Authentication
$UseIntegratedSecurity = $true


# ============================================================
# REPORT TABLES
# ============================================================

$Reports = @(
    @{
        TableName = "PC_101_DuplicateFileOccurences"
        FileName  = "PC_101_DuplicateFileOccurences.csv"
    },
    @{
        TableName = "PC_103_Files_OutsideInScopeFolders"
        FileName  = "PC_103_Files_OutsideInScopeFolders.csv"
    },
    @{
        TableName = "PC_104_Files_WhiteSpacesInName"
        FileName  = "PC_104_Files_WhiteSpacesInName.csv"
    },
    @{
        TableName = "PC_105_SWPart_ItemID_Inconsistencies"
        FileName  = "PC_105_SWPart_ItemID_Inconsistencies.csv"
    },
    @{
        TableName = "PC_106_SWPart_MultipleItemIDs"
        FileName  = "PC_106_SWPart_MultipleItemIDs.csv"
    },
    @{
        TableName = "PC_107_Files_SameName_DifferentExtension"
        FileName  = "PC_107_Files_SameName_DifferentExtension.csv"
    },
    @{
        TableName = "PC_108_Child_DistinctMaterials"
        FileName  = "PC_108_Child_DistinctMaterials.csv"
    }
	@{
        TableName = "PC_109_Child_DistinctPoids"
        FileName  = "PC_109_Child_DistinctPoids.csv"
    }
	@{
        TableName = "PC_110_ModelNotConsumed"
        FileName  = "PC_110_ModelNotConsumed.csv"
    }
	@{
        TableName = "PC_111_Child_NonNumericPoids"
        FileName  = "PC_111_Child_NonNumericPoids.csv"
    }
	@{
        TableName = "PC_113_MatchingNonCADFiles"
        FileName  = "PC_113_MatchingNonCADFiles.csv"
    }
	@{
        TableName = "PC_114_NotMatchingNonCADFiles"
        FileName  = "PC_114_NotMatchingNonCADFiles.csv"
    }
	@{
        TableName = "PC_115_UserAndDateReport"
        FileName  = "PC_115_UserAndDateReport.csv"
    }
)


# ============================================================
# CREATE OUTPUT DIRECTORY
# ============================================================

if (-not (Test-Path -LiteralPath $OutputFolder))
{
    New-Item `
        -ItemType Directory `
        -Path $OutputFolder `
        -Force |
        Out-Null
}


# ============================================================
# SQL CONNECTION
# ============================================================

if ($UseIntegratedSecurity)
{
    $ConnectionString =
        "Server=$SqlServer;" +
        "Database=$Database;" +
        "Integrated Security=True;" +
        "TrustServerCertificate=True;"
}
else
{
    throw "SQL authentication is not configured in this script."
}


# ============================================================
# LOAD SQL CLIENT
# ============================================================

try
{
    Add-Type -AssemblyName System.Data
}
catch
{
    Write-Host "Unable to load System.Data." -ForegroundColor Red
    exit 1
}


# ============================================================
# CSV ESCAPE FUNCTION
# ============================================================

function ConvertTo-CsvField
{
    param
    (
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value)
    {
        return '""'
    }

    $Text = [string]$Value

    # Escape embedded double quotes.
    #
    # CSV standard:
    #
    # " becomes ""
    #
    $Text = $Text.Replace('"', '""')

    # Always quote every field.
    #
    # This makes the CSV completely safe for:
    #   commas
    #   pipes
    #   quotes
    #   spaces
    #   paths
    #
    return '"' + $Text + '"'
}


# ============================================================
# EXPORT ONE TABLE
# ============================================================

function Export-SqlTableToCsv
{
    param
    (
        [string]$TableName,

        [string]$OutputFile
    )

    Write-Host ""
    Write-Host "Exporting: $TableName" -ForegroundColor Cyan

    $Connection = New-Object System.Data.SqlClient.SqlConnection
    $Connection.ConnectionString = $ConnectionString

    $Command = $Connection.CreateCommand()

    # Table names come from the hard-coded Reports list above.
    $Command.CommandText = @"
SELECT *
FROM dbo.[$TableName];
"@

    $Command.CommandTimeout = 0

    $DataTable = New-Object System.Data.DataTable

    try
    {
        $Connection.Open()

        $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $Adapter.SelectCommand = $Command

        [void]$Adapter.Fill($DataTable)

        $Connection.Close()
    }
    catch
    {
        if ($Connection.State -eq `
            [System.Data.ConnectionState]::Open)
        {
            $Connection.Close()
        }

        Write-Host ""
        Write-Host "ERROR exporting $TableName" `
            -ForegroundColor Red

        Write-Host $_.Exception.Message `
            -ForegroundColor Red

        return $false
    }


    # ========================================================
    # CREATE CSV CONTENT
    # ========================================================

    $StringBuilder =
        New-Object System.Text.StringBuilder


    # ========================================================
    # HEADER
    # ========================================================

    $HeaderValues = @()

    foreach ($Column in $DataTable.Columns)
    {
        $HeaderValues +=
            ConvertTo-CsvField $Column.ColumnName
    }

    [void]$StringBuilder.AppendLine(
        ($HeaderValues -join ',')
    )


    # ========================================================
    # DATA
    # ========================================================

    foreach ($Row in $DataTable.Rows)
    {
        $Values = @()

        foreach ($Column in $DataTable.Columns)
        {
            $Values +=
                ConvertTo-CsvField $Row[$Column.ColumnName]
        }

        [void]$StringBuilder.AppendLine(
            ($Values -join ',')
        )
    }


    # ========================================================
    # WRITE UTF-8 WITH BOM
    #
    # UTF-8 BOM helps Excel correctly identify Unicode data.
    # ========================================================

    try
    {
        $Utf8Bom =
            New-Object System.Text.UTF8Encoding(
                $true
            )

        [System.IO.File]::WriteAllText(
            $OutputFile,
            $StringBuilder.ToString(),
            $Utf8Bom
        )
    }
    catch
    {
        Write-Host ""
        Write-Host "ERROR writing file:" `
            -ForegroundColor Red

        Write-Host $OutputFile `
            -ForegroundColor Red

        Write-Host $_.Exception.Message `
            -ForegroundColor Red

        return $false
    }


    # ========================================================
    # RESULT
    # ========================================================

    $RowCount = $DataTable.Rows.Count

    $FileInfo = Get-Item -LiteralPath $OutputFile

    Write-Host "  Rows    : $RowCount" `
        -ForegroundColor Green

    Write-Host "  File    : $OutputFile" `
        -ForegroundColor Green

    Write-Host "  Size    : $($FileInfo.Length) bytes" `
        -ForegroundColor Green

    return $true
}


# ============================================================
# START
# ============================================================

Write-Host ""
Write-Host "============================================================"
Write-Host " SQL ANALYSIS REPORT EXPORT"
Write-Host "============================================================"
Write-Host ""

Write-Host "SQL Server : $SqlServer"
Write-Host "Database   : $Database"
Write-Host "Output     : $OutputFolder"

Write-Host ""
Write-Host "Reports to export:"
Write-Host "  1. PC_101_DuplicateFileOccurences"
Write-Host "  2. PC_103_Files_OutsideInScopeFolders"
Write-Host "  3. PC_104_Files_WhiteSpacesInName"
Write-Host "  4. PC_105_SWPart_ItemID_Inconsistencies"
Write-Host "  5. PC_106_SWPart_MultipleItemIDs"
Write-Host "  6. PC_107_Files_SameName_DifferentExtension"
Write-Host "  7. PC_108_Child_DistinctMaterials"
Write-Host "  8. PC_109_Child_DistinctPoids"
Write-Host "  9. PC_110_ModelNotConsumed"
Write-Host "  10. PC_111_Child_NonNumericPoids"

Write-Host ""
Write-Host "CSV format:"
Write-Host "  Column delimiter : ,"
Write-Host "  Text qualifier   : """""
Write-Host "  Multi-value      : |"
Write-Host "  Encoding         : UTF-8 BOM"
Write-Host ""


# ============================================================
# TEST CONNECTION
# ============================================================

Write-Host "Testing SQL connection..." `
    -ForegroundColor Yellow

$TestConnection =
    New-Object System.Data.SqlClient.SqlConnection

$TestConnection.ConnectionString =
    $ConnectionString

try
{
    $TestConnection.Open()

    Write-Host "SQL connection successful." `
        -ForegroundColor Green

    $TestConnection.Close()
}
catch
{
    Write-Host ""
    Write-Host "SQL connection FAILED." `
        -ForegroundColor Red

    Write-Host $_.Exception.Message `
        -ForegroundColor Red

    exit 1
}


# ============================================================
# EXPORT REPORTS
# ============================================================

$SuccessCount = 0
$FailureCount = 0


foreach ($Report in $Reports)
{
    $OutputFile = Join-Path `
        $OutputFolder `
        $Report.FileName

    $Success =
        Export-SqlTableToCsv `
            -TableName $Report.TableName `
            -OutputFile $OutputFile

    if ($Success)
    {
        $SuccessCount++
    }
    else
    {
        $FailureCount++
    }
}


# ============================================================
# FINAL SUMMARY
# ============================================================

Write-Host ""
Write-Host "============================================================"
Write-Host " EXPORT COMPLETE"
Write-Host "============================================================"

Write-Host ""
Write-Host "Successful : $SuccessCount" `
    -ForegroundColor Green

Write-Host "Failed     : $FailureCount" `
    -ForegroundColor $(if ($FailureCount -eq 0)
                        {
                            "Green"
                        }
                        else
                        {
                            "Red"
                        })

Write-Host ""
Write-Host "Output folder:"
Write-Host $OutputFolder `
    -ForegroundColor Cyan

Write-Host ""

if ($FailureCount -eq 0)
{
    Write-Host "All reports exported successfully." `
        -ForegroundColor Green
}
else
{
    Write-Host "One or more reports failed." `
        -ForegroundColor Red
}

Write-Host ""

