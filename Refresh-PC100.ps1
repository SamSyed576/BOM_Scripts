#requires -Version 5.1

<#
.SYNOPSIS
    Refreshes dbo.PC_100_ConsolidatedBOM from Consolidated_BOM.csv
    through dbo.PC_100_ConsolidatedBOM_Staging.

.DESCRIPTION
    CSV is the source of truth.

    STAGING:
      - Must contain exactly the CSV columns.
      - All columns are NVARCHAR(MAX).
      - No SQL-owned columns.

    PRODUCTION:
      - Must contain all CSV columns.
      - All columns are NVARCHAR(MAX).
      - ItemID is the only permitted SQL-owned column.
      - ItemID is optional.
      - If ItemID exists, it is populated after the CSV data is loaded.

    Data loading:
      CSV -> Staging using streaming SqlBulkCopy.
      Staging -> Production using a transactional INSERT.

    IMPORTANT:
      - Every CSV/Production/Staging column is treated as text.
      - No date/time or numeric conversion is performed.
      - Blank CSV values are loaded as NULL.
      - Invalid typed values cannot cause a conversion failure because
        all Production columns are NVARCHAR(MAX).

    Tables are never dropped/recreated.

    _SourceFile is treated as a normal CSV column and is not renamed.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"


# ============================================================
# CONFIGURATION
# ============================================================

$ServerName = "localhost"
$DatabaseName = "StagingDB"

$ProductionTable = "dbo.PC_100_ConsolidatedBOM"
$StagingTable = "dbo.PC_100_ConsolidatedBOM_Staging"

$CsvPath = "D:\Princecraft\BOM_Consolidation_Utility\Output\Consolidated_BOM.csv"

$ProgressInterval = 5000
$BulkBatchSize = 10000

# Production may contain ItemID even though it is not in the CSV.
$ProductionSqlOwnedColumns = @("ItemID")

# Staging must contain only CSV columns.
$StagingSqlOwnedColumns = @()


# ============================================================
# OUTPUT HELPERS
# ============================================================

function Write-Info {
    param(
        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$Message
    )

    Write-Host $Message
}


function Write-OK {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Message
    )

    Write-Host "[OK] $Message" -ForegroundColor Green
}


function Write-Warn {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Message
    )

    Write-Host "[WARN] $Message" -ForegroundColor Yellow
}


function Write-Fail {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Message
    )

    Write-Host "[FAIL] $Message" -ForegroundColor Red
}


# ============================================================
# SQL IDENTIFIER HELPERS
# ============================================================

function Quote-SqlIdentifier {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    return "[" + $Name.Replace("]", "]]") + "]"
}


function Split-TableName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TableName
    )

    $parts = $TableName.Split(".")

    if ($parts.Count -eq 1) {
        return [PSCustomObject]@{
            Schema = "dbo"
            Table  = $parts[0]
        }
    }

    if ($parts.Count -eq 2) {
        return [PSCustomObject]@{
            Schema = $parts[0]
            Table  = $parts[1]
        }
    }

    throw "Unsupported table name '$TableName'. Expected TableName or Schema.TableName."
}


function Get-QualifiedTableName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TableName
    )

    $parts = Split-TableName -TableName $TableName

    return (
        (Quote-SqlIdentifier $parts.Schema) +
        "." +
        (Quote-SqlIdentifier $parts.Table)
    )
}


# ============================================================
# COLUMN NAME NORMALIZATION
# ============================================================

function Normalize-ColumnNameForComparison {
    param(
        [AllowNull()]
        [string]$Name
    )

    if ($null -eq $Name) {
        return ""
    }

    return $Name.Trim().ToLowerInvariant()
}


# ============================================================
# SQL EXECUTION
# ============================================================

function Invoke-SqlNonQuery {
    param(
        [Parameter(Mandatory = $true)]
        [System.Data.SqlClient.SqlConnection]$Connection,

        [Parameter(Mandatory = $true)]
        [string]$Sql,

        [Parameter(Mandatory = $false)]
        [System.Data.SqlClient.SqlTransaction]$Transaction
    )

    $command = $null

    try {
        $command = $Connection.CreateCommand()
        $command.CommandText = $Sql
        $command.CommandTimeout = 0

        if ($null -ne $Transaction) {
            $command.Transaction = $Transaction
        }

        return $command.ExecuteNonQuery()
    }
    finally {
        if ($null -ne $command) {
            try {
                $command.Dispose()
            }
            catch {
            }
        }
    }
}


function Invoke-SqlScalar {
    param(
        [Parameter(Mandatory = $true)]
        [System.Data.SqlClient.SqlConnection]$Connection,

        [Parameter(Mandatory = $true)]
        [string]$Sql,

        [Parameter(Mandatory = $false)]
        [System.Data.SqlClient.SqlTransaction]$Transaction
    )

    $command = $null

    try {
        $command = $Connection.CreateCommand()
        $command.CommandText = $Sql
        $command.CommandTimeout = 0

        if ($null -ne $Transaction) {
            $command.Transaction = $Transaction
        }

        return $command.ExecuteScalar()
    }
    finally {
        if ($null -ne $command) {
            try {
                $command.Dispose()
            }
            catch {
            }
        }
    }
}


# ============================================================
# GET SQL TABLE SCHEMA
# ============================================================

function Get-TableSchema {
    param(
        [Parameter(Mandatory = $true)]
        [System.Data.SqlClient.SqlConnection]$Connection,

        [Parameter(Mandatory = $true)]
        [string]$TableName
    )

    $parts = Split-TableName -TableName $TableName

    Write-Info "Reading existing schema for $TableName..."

    $sql = @"
SELECT
    c.column_id AS ColumnID,
    c.name AS SqlColumnName,
    ty.name AS DataType,
    c.max_length AS MaxLength,
    c.precision AS PrecisionValue,
    c.scale AS ScaleValue,
    c.is_nullable AS IsNullable,
    c.is_identity AS IsIdentity,
    c.is_computed AS IsComputed,
    dc.definition AS DefaultDefinition
FROM sys.columns c
INNER JOIN sys.tables t
    ON c.object_id = t.object_id
INNER JOIN sys.schemas s
    ON t.schema_id = s.schema_id
INNER JOIN sys.types ty
    ON c.user_type_id = ty.user_type_id
LEFT JOIN sys.default_constraints dc
    ON c.default_object_id = dc.object_id
WHERE s.name = @SchemaName
  AND t.name = @TableName
ORDER BY c.column_id;
"@

    $command = $null
    $adapter = $null
    $dataTable = $null

    try {
        $command = $Connection.CreateCommand()
        $command.CommandText = $sql
        $command.CommandTimeout = 0

        $parameter = $command.Parameters.Add(
            "@SchemaName",
            [System.Data.SqlDbType]::NVarChar,
            128
        )
        $parameter.Value = $parts.Schema

        $parameter = $command.Parameters.Add(
            "@TableName",
            [System.Data.SqlDbType]::NVarChar,
            128
        )
        $parameter.Value = $parts.Table

        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($command)
        $dataTable = New-Object System.Data.DataTable

        [void]$adapter.Fill($dataTable)

        $results = @()

        foreach ($dataRow in $dataTable.Rows) {

            $defaultDefinition = $null

            if (-not $dataRow.IsNull("DefaultDefinition")) {
                $defaultDefinition = [string]$dataRow["DefaultDefinition"]
            }

            $row = [PSCustomObject]@{
                ColumnID          = [int]$dataRow["ColumnID"]
                SqlColumnName     = [string]$dataRow["SqlColumnName"]
                DataType          = [string]$dataRow["DataType"]
                MaxLength         = [int]$dataRow["MaxLength"]
                PrecisionValue    = [int]$dataRow["PrecisionValue"]
                ScaleValue        = [int]$dataRow["ScaleValue"]
                IsNullable        = [bool]$dataRow["IsNullable"]
                IsIdentity        = [bool]$dataRow["IsIdentity"]
                IsComputed        = [bool]$dataRow["IsComputed"]
                DefaultDefinition = $defaultDefinition
            }

            $results += $row
        }

        return $results
    }
    finally {

        if ($null -ne $adapter) {
            try {
                $adapter.Dispose()
            }
            catch {
            }
        }

        if ($null -ne $command) {
            try {
                $command.Dispose()
            }
            catch {
            }
        }

        if ($null -ne $dataTable) {
            try {
                $dataTable.Dispose()
            }
            catch {
            }
        }
    }
}


# ============================================================
# ADD MISSING CSV COLUMNS
#
# Any CSV column missing from the SQL table is added as
# NVARCHAR(MAX). Existing columns are left intact here and
# converted later by Convert-AllColumnsToNvarcharMax.
# ============================================================

function Add-MissingCsvColumns {
    param(
        [Parameter(Mandatory = $true)]
        [System.Data.SqlClient.SqlConnection]$Connection,

        [Parameter(Mandatory = $true)]
        [string]$TableName,

        [Parameter(Mandatory = $true)]
        [string[]]$CsvHeaders
    )

    Write-Info "Checking for CSV columns missing from $TableName..."

    $schema = @(Get-TableSchema `
        -Connection $Connection `
        -TableName $TableName)

    $existingLookup = @{}

    foreach ($column in $schema) {

        $normalized = Normalize-ColumnNameForComparison `
            $column.SqlColumnName

        if ($existingLookup.ContainsKey($normalized)) {
            throw (
                "Duplicate SQL columns found after normalization in " +
                "${TableName}: '$($existingLookup[$normalized].SqlColumnName)' " +
                "and '$($column.SqlColumnName)'."
            )
        }

        $existingLookup[$normalized] = $column
    }

    $added = 0

    foreach ($csvHeader in $CsvHeaders) {

        $normalized = Normalize-ColumnNameForComparison $csvHeader

        if (-not $existingLookup.ContainsKey($normalized)) {

            $sql = @"
ALTER TABLE $(Get-QualifiedTableName $TableName)
ADD $(Quote-SqlIdentifier $csvHeader) NVARCHAR(MAX) NULL;
"@

            Write-Info "Adding missing CSV column '$csvHeader' to $TableName..."

            [void](Invoke-SqlNonQuery `
                -Connection $Connection `
                -Sql $sql)

            $added++
        }
    }

    if ($added -eq 0) {
        Write-OK "No missing CSV columns found in $TableName."
    }
    else {
        Write-OK "Added $added missing CSV column(s) to $TableName."
    }
}


# ============================================================
# REMOVE OBSOLETE SQL COLUMNS
#
# Removes columns that are not in the CSV and are not explicitly
# allowed SQL-owned columns.
# ============================================================

function Remove-ObsoleteSqlColumns {
    param(
        [Parameter(Mandatory = $true)]
        [System.Data.SqlClient.SqlConnection]$Connection,

        [Parameter(Mandatory = $true)]
        [string]$TableName,

        [Parameter(Mandatory = $true)]
        [string[]]$CsvHeaders,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]$SqlOwnedColumns
    )

    Write-Info "Checking for obsolete columns in $TableName..."

    $schema = @(Get-TableSchema `
        -Connection $Connection `
        -TableName $TableName)

    $csvLookup = @{}

    foreach ($header in $CsvHeaders) {

        $normalized = Normalize-ColumnNameForComparison $header

        if ($csvLookup.ContainsKey($normalized)) {
            throw "Duplicate CSV column detected: '$header'."
        }

        $csvLookup[$normalized] = $header
    }

    $ownedLookup = @{}

    foreach ($owned in $SqlOwnedColumns) {

        $normalized = Normalize-ColumnNameForComparison $owned

        if ($normalized -ne "") {
            $ownedLookup[$normalized] = $true
        }
    }

    $removed = 0

    foreach ($column in $schema) {

        $normalizedSqlName = Normalize-ColumnNameForComparison `
            $column.SqlColumnName

        $isCsvColumn = $csvLookup.ContainsKey($normalizedSqlName)
        $isOwnedColumn = $ownedLookup.ContainsKey($normalizedSqlName)

        if (-not $isCsvColumn -and -not $isOwnedColumn) {

            if ($column.IsIdentity -or $column.IsComputed) {
                throw (
                    "Obsolete column '$($column.SqlColumnName)' in " +
                    "$TableName is identity/computed and cannot be " +
                    "automatically removed."
                )
            }

            Write-Warn (
                "Removing obsolete column '$($column.SqlColumnName)' " +
                "from $TableName."
            )

            $sql = @"
ALTER TABLE $(Get-QualifiedTableName $TableName)
DROP COLUMN $(Quote-SqlIdentifier $column.SqlColumnName);
"@

            [void](Invoke-SqlNonQuery `
                -Connection $Connection `
                -Sql $sql)

            $removed++
        }
    }

    if ($removed -eq 0) {
        Write-OK "No obsolete columns found in $TableName."
    }
    else {
        Write-OK "Removed $removed obsolete column(s) from $TableName."
    }
}


# ============================================================
# CONVERT ALL COLUMNS TO NVARCHAR(MAX)
#
# Every existing column in the specified table must be
# NVARCHAR(MAX).
#
# Identity and computed columns cannot be automatically
# converted and therefore cause the refresh to stop.
#
# Existing NULL / NOT NULL behavior is preserved.
# ============================================================

function Convert-AllColumnsToNvarcharMax {
    param(
        [Parameter(Mandatory = $true)]
        [System.Data.SqlClient.SqlConnection]$Connection,

        [Parameter(Mandatory = $true)]
        [string]$TableName
    )

    Write-Info ""
    Write-Info "Converting all columns in $TableName to NVARCHAR(MAX)..."

    $schema = @(Get-TableSchema `
        -Connection $Connection `
        -TableName $TableName)

    $converted = 0
    $alreadyCorrect = 0

    foreach ($column in $schema) {

        if ($column.IsIdentity) {
            throw (
                "Column '$($column.SqlColumnName)' in '$TableName' " +
                "is an IDENTITY column. It cannot be converted to " +
                "NVARCHAR(MAX) automatically."
            )
        }

        if ($column.IsComputed) {
            throw (
                "Column '$($column.SqlColumnName)' in '$TableName' " +
                "is a computed column. It cannot be converted to " +
                "NVARCHAR(MAX) automatically."
            )
        }

        if (
            $column.DataType.ToLowerInvariant() -eq "nvarchar" -and
            $column.MaxLength -eq -1
        ) {
            $alreadyCorrect++
            continue
        }

        $nullability = "NULL"

        if (-not $column.IsNullable) {
            $nullability = "NOT NULL"
        }

        Write-Info (
            "Converting '$($column.SqlColumnName)' " +
            "($($column.DataType)) -> NVARCHAR(MAX) $nullability"
        )

        $sql = @"
ALTER TABLE $(Get-QualifiedTableName $TableName)
ALTER COLUMN $(Quote-SqlIdentifier $column.SqlColumnName) NVARCHAR(MAX) $nullability;
"@

        [void](Invoke-SqlNonQuery `
            -Connection $Connection `
            -Sql $sql)

        $converted++
    }

    Write-OK (
        "${TableName}: $converted column(s) converted, " +
        "$alreadyCorrect already NVARCHAR(MAX)."
    )
}


# ============================================================
# VALIDATE TABLE SCHEMA
#
# Confirms:
#   - every CSV column exists
#   - every column is NVARCHAR(MAX)
#   - no CSV column is computed/identity
#   - no unexpected SQL-only columns remain
# ============================================================

function Validate-TableSchema {
    param(
        [Parameter(Mandatory = $true)]
        [System.Data.SqlClient.SqlConnection]$Connection,

        [Parameter(Mandatory = $true)]
        [string]$TableName,

        [Parameter(Mandatory = $true)]
        [string[]]$CsvHeaders,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]$SqlOwnedColumns
    )

    Write-Info "Validating final schema for $TableName..."

    $schema = @(Get-TableSchema `
        -Connection $Connection `
        -TableName $TableName)

    $schemaLookup = @{}

    foreach ($column in $schema) {

        $normalized = Normalize-ColumnNameForComparison `
            $column.SqlColumnName

        if ($schemaLookup.ContainsKey($normalized)) {
            throw (
                "Duplicate SQL columns found after normalization in " +
                "${TableName}: '$($schemaLookup[$normalized].SqlColumnName)' " +
                "and '$($column.SqlColumnName)'."
            )
        }

        $schemaLookup[$normalized] = $column
    }

    $csvLookup = @{}

    foreach ($header in $CsvHeaders) {

        $normalized = Normalize-ColumnNameForComparison $header

        if ($csvLookup.ContainsKey($normalized)) {
            throw "Duplicate CSV header detected: '$header'."
        }

        $csvLookup[$normalized] = $header

        if (-not $schemaLookup.ContainsKey($normalized)) {
            throw (
                "CSV column '$header' is missing from table '$TableName'."
            )
        }

        $sqlColumn = $schemaLookup[$normalized]

        if ($sqlColumn.IsComputed) {
            throw (
                "CSV column '$header' maps to computed SQL column " +
                "'$($sqlColumn.SqlColumnName)' in '$TableName'."
            )
        }

        if ($sqlColumn.IsIdentity) {
            throw (
                "CSV column '$header' maps to identity SQL column " +
                "'$($sqlColumn.SqlColumnName)' in '$TableName'."
            )
        }

        if (
            $sqlColumn.DataType.ToLowerInvariant() -ne "nvarchar" -or
            $sqlColumn.MaxLength -ne -1
        ) {
            throw (
                "Column '$($sqlColumn.SqlColumnName)' in '$TableName' " +
                "is not NVARCHAR(MAX). Actual type: " +
                "$($sqlColumn.DataType), max_length=$($sqlColumn.MaxLength)."
            )
        }
    }

    $ownedLookup = @{}

    foreach ($owned in $SqlOwnedColumns) {

        $normalized = Normalize-ColumnNameForComparison $owned

        if ($normalized -ne "") {
            $ownedLookup[$normalized] = $true
        }
    }

    foreach ($column in $schema) {

        $normalized = Normalize-ColumnNameForComparison `
            $column.SqlColumnName

        $isCsvColumn = $csvLookup.ContainsKey($normalized)
        $isOwnedColumn = $ownedLookup.ContainsKey($normalized)

        if (-not $isCsvColumn -and -not $isOwnedColumn) {
            throw (
                "Unexpected SQL column '$($column.SqlColumnName)' " +
                "remains in '$TableName'."
            )
        }

        if (
            $column.DataType.ToLowerInvariant() -ne "nvarchar" -or
            $column.MaxLength -ne -1
        ) {
            throw (
                "Column '$($column.SqlColumnName)' in '$TableName' " +
                "is not NVARCHAR(MAX). Actual type: " +
                "$($column.DataType), max_length=$($column.MaxLength)."
            )
        }
    }

    Write-OK (
        "$TableName schema validated: " +
        "$($CsvHeaders.Count) CSV column(s), " +
        "$($schema.Count) total SQL column(s), " +
        "all NVARCHAR(MAX)."
    )

    return $schema
}


# ============================================================
# SYNCHRONIZE TABLE SCHEMA
# ============================================================

function Sync-TableSchema {
    param(
        [Parameter(Mandatory = $true)]
        [System.Data.SqlClient.SqlConnection]$Connection,

        [Parameter(Mandatory = $true)]
        [string]$TableName,

        [Parameter(Mandatory = $true)]
        [string[]]$CsvHeaders,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]$SqlOwnedColumns
    )

    Write-Info ""
    Write-Info "Synchronizing schema: $TableName"

    Add-MissingCsvColumns `
        -Connection $Connection `
        -TableName $TableName `
        -CsvHeaders $CsvHeaders

    Remove-ObsoleteSqlColumns `
        -Connection $Connection `
        -TableName $TableName `
        -CsvHeaders $CsvHeaders `
        -SqlOwnedColumns $SqlOwnedColumns

    Convert-AllColumnsToNvarcharMax `
        -Connection $Connection `
        -TableName $TableName

    $finalSchema = @(Validate-TableSchema `
        -Connection $Connection `
        -TableName $TableName `
        -CsvHeaders $CsvHeaders `
        -SqlOwnedColumns $SqlOwnedColumns)

    return $finalSchema
}


# ============================================================
# READ CSV HEADER
# ============================================================

function Read-CsvHeaders {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    Write-Info "Reading CSV header..."

    Add-Type -AssemblyName Microsoft.VisualBasic

    $parser = $null

    try {

        $parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($Path)

        $parser.TextFieldType = `
            [Microsoft.VisualBasic.FileIO.FieldType]::Delimited

        $parser.SetDelimiters(",")

        $parser.HasFieldsEnclosedInQuotes = $true
        $parser.TrimWhiteSpace = $false

        $headers = $parser.ReadFields()

        if ($null -eq $headers) {
            throw "CSV header could not be read."
        }

        if ($headers.Count -eq 0) {
            throw "CSV contains no columns."
        }

        if ($headers[0].Length -gt 0 -and `
            [int][char]$headers[0][0] -eq 0xFEFF) {

            $headers[0] = $headers[0].Substring(1)
        }

        $lookup = @{}

        foreach ($header in $headers) {

            if ($null -eq $header) {
                throw "CSV contains a NULL header."
            }

            if ([string]::IsNullOrWhiteSpace($header)) {
                throw "CSV contains a blank header."
            }

            $normalized = Normalize-ColumnNameForComparison $header

            if ($lookup.ContainsKey($normalized)) {
                throw (
                    "Duplicate CSV header detected after normalization: " +
                    "'$header' and '$($lookup[$normalized])'."
                )
            }

            $lookup[$normalized] = $header
        }

        Write-OK "CSV contains $($headers.Count) columns."

        return [string[]]$headers
    }
    finally {

        if ($null -ne $parser) {

            try {
                $parser.Close()
            }
            catch {
            }

            try {
                $parser.Dispose()
            }
            catch {
            }
        }
    }
}


# ============================================================
# INITIALIZE C# STREAMING CSV READER
# ============================================================

function Initialize-CsvStreamingReaderType {

    if ($null -ne ("PrincecraftCsvStreamingReader" -as [type])) {
        return
    }

    $source = @"
using System;
using System.Data;
using System.Collections.Generic;
using Microsoft.VisualBasic.FileIO;

public sealed class PrincecraftCsvStreamingReader : IDataReader
{
    private readonly TextFieldParser _parser;
    private readonly string[] _headers;
    private readonly Dictionary<string, int> _ordinalLookup;
    private readonly int _topLineOrdinal;

    private string[] _currentValues;
    private bool _closed;
    private long _recordsRead;

    public PrincecraftCsvStreamingReader(
        string csvPath,
        string[] headers,
        int topLineOrdinal)
    {
        if (csvPath == null)
            throw new ArgumentNullException("csvPath");

        if (headers == null)
            throw new ArgumentNullException("headers");

        _headers = headers;
        _topLineOrdinal = topLineOrdinal;

        _ordinalLookup =
            new Dictionary<string, int>(
                StringComparer.OrdinalIgnoreCase);

        for (int i = 0; i < _headers.Length; i++)
        {
            if (_ordinalLookup.ContainsKey(_headers[i]))
                throw new InvalidOperationException(
                    "Duplicate CSV header: " + _headers[i]);

            _ordinalLookup.Add(_headers[i], i);
        }

        _parser = new TextFieldParser(csvPath);
        _parser.TextFieldType = FieldType.Delimited;
        _parser.SetDelimiters(",");
        _parser.HasFieldsEnclosedInQuotes = true;
        _parser.TrimWhiteSpace = false;

        string[] actualHeader = _parser.ReadFields();

        if (actualHeader == null)
            throw new InvalidOperationException(
                "CSV header could not be read.");

        if (actualHeader.Length != _headers.Length)
            throw new InvalidOperationException(
                "CSV header count does not match expected header count.");

        for (int i = 0; i < actualHeader.Length; i++)
        {
            string actual = actualHeader[i];

            if (i == 0 &&
                actual.Length > 0 &&
                actual[0] == '\uFEFF')
            {
                actual = actual.Substring(1);
            }

            if (!String.Equals(
                    actual,
                    _headers[i],
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "CSV header mismatch at column " +
                    (i + 1) +
                    ". Expected '" +
                    _headers[i] +
                    "' but found '" +
                    actual +
                    "'.");
            }
        }

        _currentValues = null;
        _closed = false;
        _recordsRead = 0;
    }

    public bool Read()
    {
        if (_closed)
            throw new ObjectDisposedException(
                "PrincecraftCsvStreamingReader");

        if (_parser.EndOfData)
        {
            _currentValues = null;
            return false;
        }

        string[] fields = _parser.ReadFields();

        if (fields == null)
        {
            _currentValues = null;
            return false;
        }

        if (fields.Length != _headers.Length)
        {
            throw new InvalidOperationException(
                "CSV row " +
                (_recordsRead + 2) +
                " contains " +
                fields.Length +
                " fields but expected " +
                _headers.Length +
                ".");
        }

        if (_topLineOrdinal >= 0 &&
            _topLineOrdinal < fields.Length)
        {
            string topLine = fields[_topLineOrdinal];

            if (String.IsNullOrWhiteSpace(topLine))
                fields[_topLineOrdinal] = "NULL";
        }

        _currentValues = fields;
        _recordsRead++;

        return true;
    }

    public int FieldCount
    {
        get { return _headers.Length; }
    }

    public string GetName(int i)
    {
        return _headers[i];
    }

    public int GetOrdinal(string name)
    {
        int ordinal;

        if (!_ordinalLookup.TryGetValue(name, out ordinal))
            throw new IndexOutOfRangeException(
                "Column not found: " + name);

        return ordinal;
    }

    public object GetValue(int i)
    {
        if (_currentValues == null)
            throw new InvalidOperationException("No current row.");

        string value = _currentValues[i];

        if (String.IsNullOrEmpty(value))
            return DBNull.Value;

        return value;
    }

    public int GetValues(object[] values)
    {
        if (_currentValues == null)
            throw new InvalidOperationException("No current row.");

        int count = Math.Min(values.Length, _headers.Length);

        for (int i = 0; i < count; i++)
            values[i] = GetValue(i);

        return count;
    }

    public bool IsDBNull(int i)
    {
        return GetValue(i) == DBNull.Value;
    }

    public Type GetFieldType(int i)
    {
        return typeof(string);
    }

    public string GetDataTypeName(int i)
    {
        return "nvarchar";
    }

    public bool NextResult()
    {
        return false;
    }

    public int Depth
    {
        get { return 0; }
    }

    public bool IsClosed
    {
        get { return _closed; }
    }

    public int RecordsAffected
    {
        get { return -1; }
    }

    public DataTable GetSchemaTable()
    {
        DataTable table = new DataTable("SchemaTable");

        table.Columns.Add("ColumnName", typeof(string));
        table.Columns.Add("ColumnOrdinal", typeof(int));
        table.Columns.Add("DataType", typeof(Type));
        table.Columns.Add("AllowDBNull", typeof(bool));

        for (int i = 0; i < _headers.Length; i++)
        {
            DataRow row = table.NewRow();

            row["ColumnName"] = _headers[i];
            row["ColumnOrdinal"] = i;
            row["DataType"] = typeof(string);
            row["AllowDBNull"] = true;

            table.Rows.Add(row);
        }

        return table;
    }

    public object this[int i]
    {
        get { return GetValue(i); }
    }

    public object this[string name]
    {
        get { return GetValue(GetOrdinal(name)); }
    }

    public void Close()
    {
        if (!_closed)
        {
            _parser.Close();
            _closed = true;
        }
    }

    public void Dispose()
    {
        Close();
    }

    public IDataReader GetData(int i)
    {
        throw new NotSupportedException();
    }

    public bool GetBoolean(int i)
    {
        return Convert.ToBoolean(GetValue(i));
    }

    public byte GetByte(int i)
    {
        return Convert.ToByte(GetValue(i));
    }

    public long GetBytes(
        int i,
        long fieldOffset,
        byte[] buffer,
        int bufferoffset,
        int length)
    {
        throw new NotSupportedException();
    }

    public char GetChar(int i)
    {
        return Convert.ToChar(GetValue(i));
    }

    public long GetChars(
        int i,
        long fieldoffset,
        char[] buffer,
        int bufferoffset,
        int length)
    {
        throw new NotSupportedException();
    }

    public Guid GetGuid(int i)
    {
        return Guid.Parse(Convert.ToString(GetValue(i)));
    }

    public short GetInt16(int i)
    {
        return Convert.ToInt16(GetValue(i));
    }

    public int GetInt32(int i)
    {
        return Convert.ToInt32(GetValue(i));
    }

    public long GetInt64(int i)
    {
        return Convert.ToInt64(GetValue(i));
    }

    public float GetFloat(int i)
    {
        return Convert.ToSingle(GetValue(i));
    }

    public double GetDouble(int i)
    {
        return Convert.ToDouble(GetValue(i));
    }

    public decimal GetDecimal(int i)
    {
        return Convert.ToDecimal(GetValue(i));
    }

    public DateTime GetDateTime(int i)
    {
        return Convert.ToDateTime(GetValue(i));
    }

    public string GetString(int i)
    {
        object value = GetValue(i);

        if (value == DBNull.Value)
            return null;

        return Convert.ToString(value);
    }
}
"@

    Add-Type `
        -TypeDefinition $source `
        -Language CSharp `
        -ReferencedAssemblies @(
            "System.dll",
            "System.Data.dll",
            "System.Xml.dll",
            "Microsoft.VisualBasic.dll"
        )
}


# ============================================================
# LOAD CSV INTO STAGING
#
# CSV is streamed directly into staging.
# SQL column mappings are name-to-name rather than ordinal-based,
# so gaps in sys.columns.column_id do not cause problems.
# ============================================================

function Load-CsvIntoStaging {
    param(
        [Parameter(Mandatory = $true)]
        [System.Data.SqlClient.SqlConnection]$Connection,

        [Parameter(Mandatory = $true)]
        [string]$TableName,

        [Parameter(Mandatory = $true)]
        [string]$CsvPath,

        [Parameter(Mandatory = $true)]
        [string[]]$Headers,

        [Parameter(Mandatory = $true)]
        [int]$ProgressInterval,

        [Parameter(Mandatory = $true)]
        [int]$BatchSize
    )

    Write-Info ""
    Write-Info "6. LOADING CSV INTO STAGING"

    Write-Info "Clearing staging table..."

    $qualifiedStaging = Get-QualifiedTableName $TableName

    $deleteSql = @"
TRUNCATE TABLE $qualifiedStaging;
"@

    [void](Invoke-SqlNonQuery `
        -Connection $Connection `
        -Sql $deleteSql)

    Write-OK "Staging table cleared."

    Initialize-CsvStreamingReaderType

    $topLineOrdinal = -1

    for ($i = 0; $i -lt $Headers.Count; $i++) {

        if ((Normalize-ColumnNameForComparison $Headers[$i]) -eq "topline") {
            $topLineOrdinal = $i
            break
        }
    }

    $reader = $null
    $bulkCopy = $null

    try {

        $reader = [PrincecraftCsvStreamingReader]::new(
            $CsvPath,
            [string[]]$Headers,
            [int]$topLineOrdinal
        )

        Write-Info "Building name-to-name column mappings..."

        $bulkCopy = New-Object System.Data.SqlClient.SqlBulkCopy(
            $Connection,
            [System.Data.SqlClient.SqlBulkCopyOptions]::TableLock,
            $null
        )

        $bulkCopy.DestinationTableName = $TableName
        $bulkCopy.BatchSize = $BatchSize
        $bulkCopy.BulkCopyTimeout = 0
        $bulkCopy.NotifyAfter = $ProgressInterval

        $bulkCopy.add_SqlRowsCopied({
            param($sender, $eventArgs)

            if (($eventArgs.RowsCopied % $ProgressInterval) -eq 0) {
                Write-Info "Rows loaded: $($eventArgs.RowsCopied)"
            }
        })

        $destinationSchema = @(Get-TableSchema `
            -Connection $Connection `
            -TableName $TableName)

        $destinationLookup = @{}

        foreach ($destinationColumn in $destinationSchema) {

            $normalized = Normalize-ColumnNameForComparison `
                $destinationColumn.SqlColumnName

            $destinationLookup[$normalized] = $destinationColumn
        }

        $mappingCount = 0

        for ($sourceOrdinal = 0; `
             $sourceOrdinal -lt $Headers.Count; `
             $sourceOrdinal++) {

            $sourceColumnName = [string]$Headers[$sourceOrdinal]

            $normalizedSourceName = `
                Normalize-ColumnNameForComparison $sourceColumnName

            if (-not $destinationLookup.ContainsKey($normalizedSourceName)) {
                throw (
                    "CSV column '$sourceColumnName' does not exist " +
                    "in staging table '$TableName'."
                )
            }

            $destinationColumn = `
                $destinationLookup[$normalizedSourceName]

            $destinationColumnName = `
                [string]$destinationColumn.SqlColumnName

            $mapping = New-Object `
                System.Data.SqlClient.SqlBulkCopyColumnMapping(
                    $sourceColumnName,
                    $destinationColumnName
                )

            [void]$bulkCopy.ColumnMappings.Add($mapping)

            $mappingCount++
        }

        if ($mappingCount -ne $Headers.Count) {
            throw (
                "SqlBulkCopy mapping count $mappingCount does not " +
                "match CSV column count $($Headers.Count)."
            )
        }

        Write-Info "SqlBulkCopy mappings created: $mappingCount"
        Write-Info "Starting streaming CSV bulk load..."

        $bulkCopy.WriteToServer($reader)

        Write-OK "CSV loaded into staging table."
    }
    finally {

        if ($null -ne $bulkCopy) {

            try {
                $bulkCopy.Close()
            }
            catch {
            }

            try {
                $bulkCopy.Dispose()
            }
            catch {
            }
        }

        if ($null -ne $reader) {

            try {
                $reader.Close()
            }
            catch {
            }

            try {
                $reader.Dispose()
            }
            catch {
            }
        }
    }
}


# ============================================================
# REFRESH PRODUCTION FROM STAGING
#
# Production is entirely NVARCHAR(MAX), so the INSERT uses
# direct column-to-column references.
#
# The entire production replacement occurs inside one
# transaction.
#
# If anything fails:
#     DELETE is rolled back.
#     INSERT is rolled back.
#     Existing Production data remains intact.
# ============================================================

function Refresh-ProductionFromStaging {
    param(
        [Parameter(Mandatory = $true)]
        [System.Data.SqlClient.SqlConnection]$Connection,

        [Parameter(Mandatory = $true)]
        [string]$ProductionTable,

        [Parameter(Mandatory = $true)]
        [string]$StagingTable,

        [Parameter(Mandatory = $true)]
        [string[]]$Headers
    )

    Write-Info ""
    Write-Info "7. REFRESHING PRODUCTION"

    Write-Info "Reading existing schema for $ProductionTable..."

    $productionSchema = @(Get-TableSchema `
        -Connection $Connection `
        -TableName $ProductionTable)

    Write-Info "Reading existing schema for $StagingTable..."

    $stagingSchema = @(Get-TableSchema `
        -Connection $Connection `
        -TableName $StagingTable)

    $productionLookup = @{}
    $stagingLookup = @{}

    foreach ($column in $productionSchema) {

        $normalized = Normalize-ColumnNameForComparison `
            $column.SqlColumnName

        $productionLookup[$normalized] = $column
    }

    foreach ($column in $stagingSchema) {

        $normalized = Normalize-ColumnNameForComparison `
            $column.SqlColumnName

        $stagingLookup[$normalized] = $column
    }

    $targetColumns = New-Object `
        System.Collections.Generic.List[string]

    $sourceExpressions = New-Object `
        System.Collections.Generic.List[string]

    foreach ($header in $Headers) {

        $normalized = Normalize-ColumnNameForComparison $header

        if (-not $productionLookup.ContainsKey($normalized)) {
            throw (
                "CSV column '$header' does not exist in Production."
            )
        }

        if (-not $stagingLookup.ContainsKey($normalized)) {
            throw (
                "CSV column '$header' does not exist in Staging."
            )
        }

        $productionColumn = $productionLookup[$normalized]
        $stagingColumn = $stagingLookup[$normalized]

        if ($productionColumn.IsIdentity -or `
            $productionColumn.IsComputed) {

            throw (
                "Production column '$($productionColumn.SqlColumnName)' " +
                "is identity/computed and cannot be loaded from staging."
            )
        }

        if (
            $productionColumn.DataType.ToLowerInvariant() -ne "nvarchar" -or
            $productionColumn.MaxLength -ne -1
        ) {
            throw (
                "Production column '$($productionColumn.SqlColumnName)' " +
                "is not NVARCHAR(MAX)."
            )
        }

        if (
            $stagingColumn.DataType.ToLowerInvariant() -ne "nvarchar" -or
            $stagingColumn.MaxLength -ne -1
        ) {
            throw (
                "Staging column '$($stagingColumn.SqlColumnName)' " +
                "is not NVARCHAR(MAX)."
            )
        }

        [void]$targetColumns.Add(
            (Quote-SqlIdentifier $productionColumn.SqlColumnName)
        )

        [void]$sourceExpressions.Add(
            ("s." + (Quote-SqlIdentifier $stagingColumn.SqlColumnName))
        )
    }

    if ($targetColumns.Count -ne $Headers.Count) {
        throw (
            "Production target column count does not match CSV column count."
        )
    }

    $itemIdProductionColumn = $null

    $itemIdNormalized = `
        Normalize-ColumnNameForComparison "ItemID"

    if ($productionLookup.ContainsKey($itemIdNormalized)) {

        $itemIdProductionColumn = `
            $productionLookup[$itemIdNormalized]

        if ($itemIdProductionColumn.IsIdentity -or `
            $itemIdProductionColumn.IsComputed) {

            throw (
                "Production ItemID column is identity/computed. " +
                "The current refresh logic expects ItemID to be writable."
            )
        }

        if (
            $itemIdProductionColumn.DataType.ToLowerInvariant() -ne "nvarchar" -or
            $itemIdProductionColumn.MaxLength -ne -1
        ) {
            throw (
                "Production ItemID column is not NVARCHAR(MAX)."
            )
        }
    }

    $targetColumnSql = `
        $targetColumns -join ",`r`n    "

    $sourceExpressionSql = `
        $sourceExpressions -join ",`r`n        "

    $qualifiedProduction = `
        Get-QualifiedTableName $ProductionTable

    $qualifiedStaging = `
        Get-QualifiedTableName $StagingTable

    $transaction = $null

    try {

        Write-Info "Beginning production refresh transaction..."

        $transaction = $Connection.BeginTransaction()

        Write-Info "Deleting existing production rows..."

        $deleteSql = @"
DELETE FROM $qualifiedProduction;
"@

        [void](Invoke-SqlNonQuery `
            -Connection $Connection `
            -Sql $deleteSql `
            -Transaction $transaction)

        Write-Info "Loading production from staging..."

        $insertSql = @"
INSERT INTO $qualifiedProduction
(
    $targetColumnSql
)
SELECT
    $sourceExpressionSql
FROM $qualifiedStaging AS s;
"@

        [void](Invoke-SqlNonQuery `
            -Connection $Connection `
            -Sql $insertSql `
            -Transaction $transaction)

        if ($null -ne $itemIdProductionColumn) {

            Write-Info "Populating Production ItemID from ChildName..."

            $itemIdSql = @"
UPDATE p
SET $(Quote-SqlIdentifier $itemIdProductionColumn.SqlColumnName) =
    CASE
        WHEN x.StartPos > 0
        THEN SUBSTRING(
            p.$(Quote-SqlIdentifier "ChildName"),
            x.StartPos,
            13
        )
        ELSE NULL
    END
FROM $qualifiedProduction p
CROSS APPLY
(
    VALUES
    (
        PATINDEX(
            '%[0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9][0-9][0-9][0-9][0-9]%',
            p.$(Quote-SqlIdentifier "ChildName")
        )
    )
) x(StartPos);
"@

            [void](Invoke-SqlNonQuery `
                -Connection $Connection `
                -Sql $itemIdSql `
                -Transaction $transaction)
        }
        else {

            Write-Info (
                "Production ItemID column not present. " +
                "Skipping ItemID population."
            )
        }

        Write-Info "Validating production row count..."

        $stagingCountSql = @"
SELECT COUNT_BIG(*)
FROM $qualifiedStaging;
"@

        $productionCountSql = @"
SELECT COUNT_BIG(*)
FROM $qualifiedProduction;
"@

        $stagingCount = [long](Invoke-SqlScalar `
            -Connection $Connection `
            -Sql $stagingCountSql `
            -Transaction $transaction)

        $productionCount = [long](Invoke-SqlScalar `
            -Connection $Connection `
            -Sql $productionCountSql `
            -Transaction $transaction)

        Write-Info "Staging rows:    $stagingCount"
        Write-Info "Production rows: $productionCount"

        if ($stagingCount -ne $productionCount) {
            throw (
                "Production row count $productionCount does not match " +
                "staging row count $stagingCount."
            )
        }

        Write-Info "Committing production transaction..."

        $transaction.Commit()

        $transaction = $null

        Write-OK "Production refresh committed successfully."
    }
    catch {

        if ($null -ne $transaction) {

            try {
                $transaction.Rollback()
                Write-Warn "Production transaction rolled back."
            }
            catch {
                Write-Warn (
                    "Production rollback failed: " +
                    $_.Exception.Message
                )
            }
        }

        throw
    }
}


# ============================================================
# FINAL VALIDATION
# ============================================================

function Validate-FinalRefresh {
    param(
        [Parameter(Mandatory = $true)]
        [System.Data.SqlClient.SqlConnection]$Connection,

        [Parameter(Mandatory = $true)]
        [string]$ProductionTable,

        [Parameter(Mandatory = $true)]
        [string]$StagingTable
    )

    Write-Info ""
    Write-Info "8. FINAL VALIDATION"

    $qualifiedProduction = `
        Get-QualifiedTableName $ProductionTable

    $qualifiedStaging = `
        Get-QualifiedTableName $StagingTable

    $stagingCountSql = @"
SELECT COUNT_BIG(*)
FROM $qualifiedStaging;
"@

    $productionCountSql = @"
SELECT COUNT_BIG(*)
FROM $qualifiedProduction;
"@

    $stagingCount = [long](Invoke-SqlScalar `
        -Connection $Connection `
        -Sql $stagingCountSql)

    $productionCount = [long](Invoke-SqlScalar `
        -Connection $Connection `
        -Sql $productionCountSql)

    Write-Info "Final staging row count:    $stagingCount"
    Write-Info "Final production row count: $productionCount"

    if ($stagingCount -ne $productionCount) {
        throw (
            "Final validation failed. Staging has $stagingCount rows " +
            "but Production has $productionCount rows."
        )
    }

    Write-OK "Final row-count validation passed."

    $productionSchema = @(Get-TableSchema `
        -Connection $Connection `
        -TableName $ProductionTable)

    $stagingSchema = @(Get-TableSchema `
        -Connection $Connection `
        -TableName $StagingTable)

    Write-Info "Final staging column count:    $($stagingSchema.Count)"
    Write-Info "Final production column count: $($productionSchema.Count)"

    if ($productionSchema.Count -lt $stagingSchema.Count) {
        throw (
            "Production has fewer columns than staging."
        )
    }

    foreach ($column in $stagingSchema) {

        if (
            $column.DataType.ToLowerInvariant() -ne "nvarchar" -or
            $column.MaxLength -ne -1
        ) {
            throw (
                "Final validation failed: staging column " +
                "'$($column.SqlColumnName)' is not NVARCHAR(MAX)."
            )
        }
    }

    foreach ($column in $productionSchema) {

        if (
            $column.DataType.ToLowerInvariant() -ne "nvarchar" -or
            $column.MaxLength -ne -1
        ) {
            throw (
                "Final validation failed: production column " +
                "'$($column.SqlColumnName)' is not NVARCHAR(MAX)."
            )
        }
    }

    Write-OK "Final NVARCHAR(MAX) schema validation passed."
}


# ============================================================
# MAIN
# ============================================================

$connection = $null

try {

    Write-Host ""
    Write-Host "============================================================"
    Write-Host "PC_100 CONSOLIDATED BOM REFRESH"
    Write-Host "============================================================"
    Write-Host "CSV:"
    Write-Host "  $CsvPath"
    Write-Host "Production:"
    Write-Host "  $ProductionTable"
    Write-Host "Staging:"
    Write-Host "  $StagingTable"
    Write-Host ""
    Write-Host "All columns: NVARCHAR(MAX)"
    Write-Host ""

    # --------------------------------------------------------
    # 1. CHECK CSV
    # --------------------------------------------------------

    Write-Info "1. CHECKING CSV"

    if (-not (Test-Path -LiteralPath $CsvPath -PathType Leaf)) {
        throw "CSV file not found: $CsvPath"
    }

    $csvItem = Get-Item -LiteralPath $CsvPath

    Write-Info (
        "CSV size: " +
        "$([math]::Round($csvItem.Length / 1MB, 2)) MB"
    )

    Write-OK "CSV file exists."

    # --------------------------------------------------------
    # 2. READ CSV HEADER
    # --------------------------------------------------------

    Write-Info ""
    Write-Info "2. READING CSV HEADER"

    $Headers = @(Read-CsvHeaders -Path $CsvPath)

    Write-Info "CSV columns: $($Headers.Count)"

    # --------------------------------------------------------
    # 3. CONNECT TO SQL SERVER
    # --------------------------------------------------------

    Write-Info ""
    Write-Info "3. CONNECTING TO SQL SERVER"

    $connectionString = (
        "Server=$ServerName;" +
        "Database=$DatabaseName;" +
        "Integrated Security=True;" +
        "TrustServerCertificate=True;"
    )

    $connection = New-Object System.Data.SqlClient.SqlConnection(
        $connectionString
    )

    $connection.Open()

    Write-OK "Connected to $ServerName / $DatabaseName."

    # --------------------------------------------------------
    # 4. SYNCHRONIZE STAGING SCHEMA
    # --------------------------------------------------------

    Write-Info ""
    Write-Info "4. SYNCHRONIZING STAGING SCHEMA"

    $stagingSchema = @(Sync-TableSchema `
        -Connection $connection `
        -TableName $StagingTable `
        -CsvHeaders $Headers `
        -SqlOwnedColumns $StagingSqlOwnedColumns)

    # --------------------------------------------------------
    # 5. SYNCHRONIZE PRODUCTION SCHEMA
    # --------------------------------------------------------

    Write-Info ""
    Write-Info "5. SYNCHRONIZING PRODUCTION SCHEMA"

    $productionSchema = @(Sync-TableSchema `
        -Connection $connection `
        -TableName $ProductionTable `
        -CsvHeaders $Headers `
        -SqlOwnedColumns $ProductionSqlOwnedColumns)

    # --------------------------------------------------------
    # 6. LOAD CSV INTO STAGING
    # --------------------------------------------------------

    Load-CsvIntoStaging `
        -Connection $connection `
        -TableName $StagingTable `
        -CsvPath $CsvPath `
        -Headers $Headers `
        -ProgressInterval $ProgressInterval `
        -BatchSize $BulkBatchSize

    # --------------------------------------------------------
    # 6A. VALIDATE STAGING
    # --------------------------------------------------------

    Write-Info ""
    Write-Info "6A. VALIDATING STAGING"

    $qualifiedStaging = `
        Get-QualifiedTableName $StagingTable

    $stagingCountSql = @"
SELECT COUNT_BIG(*)
FROM $qualifiedStaging;
"@

    $stagingCount = [long](Invoke-SqlScalar `
        -Connection $connection `
        -Sql $stagingCountSql)

    Write-Info "Staging rows loaded: $stagingCount"

    if ($stagingCount -le 0) {
        throw (
            "Staging contains zero rows. Production refresh aborted."
        )
    }

    Write-OK "Staging validation passed."

    # --------------------------------------------------------
    # 7. REFRESH PRODUCTION
    # --------------------------------------------------------

    Refresh-ProductionFromStaging `
        -Connection $connection `
        -ProductionTable $ProductionTable `
        -StagingTable $StagingTable `
        -Headers $Headers

    # --------------------------------------------------------
    # 8. FINAL VALIDATION
    # --------------------------------------------------------

    Validate-FinalRefresh `
        -Connection $connection `
        -ProductionTable $ProductionTable `
        -StagingTable $StagingTable

    Write-Host ""
    Write-Host "============================================================"
    Write-Host "PC_100 REFRESH COMPLETE"
    Write-Host "============================================================"
    Write-Host ""

}
catch {

    Write-Host ""
    Write-Fail $_.Exception.Message
    Write-Host ""

    if ($_.ScriptStackTrace) {
        Write-Host "Script stack trace:" -ForegroundColor DarkGray
        Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
    }

    exit 1
}
finally {

    if ($null -ne $connection) {

        if ($connection.State -ne `
            [System.Data.ConnectionState]::Closed) {

            try {
                $connection.Close()
            }
            catch {
            }
        }

        try {
            $connection.Dispose()
        }
        catch {
        }
    }
}