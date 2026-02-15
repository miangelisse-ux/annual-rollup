# Function: Open-Excel
# Creates a hidden Excel COM object to manipulate workbooks programmatically
function Open-Excel {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false           # Keep Excel hidden
    $excel.DisplayAlerts = $false     # Disable prompts like "Save As"
    return $excel
}

# Function: Get-Headers
# Maps header names in the Excel sheet to column numbers
function Get-Headers($vals, $cols) {
    $header = @{}
    for ($c=1; $c -le $cols; $c++) {
        $name = [string]$vals[1,$c]
        if ($name -and -not $header.ContainsKey($name)) { 
            $header[$name] = $c 
        }
    }
    return $header
}

# Function: Write-Table
# Writes output data to a new Excel workbook with formatting
# Parameters:
#   $outPath  -> full path for new file
#   $headers  -> array of column headers
#   $rows     -> array of data rows
#   $numberCols -> columns to format as numeric
function Write-Table($outPath, $headers, $rows, $numberCols) {
    $excel = Open-Excel
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)

    # Write headers
    for ($i=0; $i -lt $headers.Count; $i++) {
        $ws.Cells.Item(1, $i+1) = $headers[$i]
    }
    $ws.Range(("A1:" + [char](65 + $headers.Count - 1) + "1")).Font.Bold = $true

    # Write data rows
    $row = 2
    foreach ($r in $rows) {
        for ($i=0; $i -lt $headers.Count; $i++) {
            $ws.Cells.Item($row, $i+1) = $r[$i]
        }
        $row++
    }

    # Format numeric columns
    foreach ($col in $numberCols) {
        $ws.Range("$col`2:$col$($row-1)").NumberFormat = '#,##0.00'
    }

    # Auto-fit columns for readability
    $ws.Columns.AutoFit() | Out-Null

    # Save and close workbook
    $wb.SaveAs($outPath)
    $wb.Close($true)

    # Release Excel COM object
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

# Prompt user for input files and output folder
$janJul = Read-Host "Enter path to Jan-Jul file"
$julDec = Read-Host "Enter path to Jul-Dec file"
$selfPay = Read-Host "Enter path to Self Pay file"
$outDir = Read-Host "Enter output folder (e.g., C:\\Users\\Owner\\Downloads)"

$files = @($janJul, $julDec, $selfPay)

# ---- Insurance Rollup ----
$excel = Open-Excel
$insuranceAgg = @{}

# Loop through Jan-Jul and Jul-Dec files to aggregate insurance data
foreach ($f in $files[0..1]) {
    $wb = $excel.Workbooks.Open($f, $null, $true)
    $ws = $wb.Worksheets.Item(1)
    $used = $ws.UsedRange
    $vals = $used.Value2
    $rows = $used.Rows.Count
    $cols = $used.Columns.Count

    # Map headers to columns
    $header = Get-Headers $vals $cols
    $colIns = $header['Primary Insurance']
    $colBilled = $header['Billed Amount ($)']
    $colPosted = $header['Posted Amount ($)']
    $colBal = $header['Current Balance ($)']

    # Process each row
    for ($r=2; $r -le $rows; $r++) {
        $insRaw = [string]$vals[$r,$colIns]
        if (-not $insRaw) { $insRaw = '(Blank)' }

        # Strip any extra text in parentheses
        $ins = $insRaw
        if ($insRaw -match '^(.*?)\s*\(') { $ins = $Matches[1].Trim() }

        $billed = $vals[$r,$colBilled]
        $posted = $vals[$r,$colPosted]
        $bal = $vals[$r,$colBal]

        # Initialize object if new insurance
        if (-not $insuranceAgg.ContainsKey($ins)) {
            $insuranceAgg[$ins] = [pscustomobject]@{Insurance=$ins; Billed=0.0; Paid=0.0; Unpaid=0.0}
        }

        # Aggregate amounts
        if ($billed) { $insuranceAgg[$ins].Billed += [double]$billed }
        if ($posted) { $insuranceAgg[$ins].Paid += [double]$posted }
        if ($bal) { $insuranceAgg[$ins].Unpaid += [double]$bal }
    }
    $wb.Close($false)
}

# ---- Self Pay Rollup ----
$wb = $excel.Workbooks.Open($files[2], $null, $true)
$ws = $wb.Worksheets.Item(1)
$used = $ws.UsedRange
$vals = $used.Value2
$rows = $used.Rows.Count
$cols = $used.Columns.Count
$header = Get-Headers $vals $cols

$ins = 'Self Pay'
if (-not $insuranceAgg.ContainsKey($ins)) {
    $insuranceAgg[$ins] = [pscustomobject]@{Insurance=$ins; Billed=0.0; Paid=0.0; Unpaid=0.0}
}

# Aggregate Self Pay amounts
for ($r=2; $r -le $rows; $r++) {
    $billed = $vals[$r,$colBilled]
    $posted = $vals[$r,$colPosted]
    $bal = $vals[$r,$colBal]
    if ($billed) { $insuranceAgg[$ins].Billed += [double]$billed }
    if ($posted) { $insuranceAgg[$ins].Paid += [double]$posted }
    if ($bal) { $insuranceAgg[$ins].Unpaid += [double]$bal }
}

$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

# Prepare rows for writing
$insuranceRows = $insuranceAgg.Values |
    Sort-Object Insurance |
    ForEach-Object { @($_.Insurance, [double]$_.Billed, [double]$_.Paid, [double]$_.Unpaid) }

# Write Insurance summary Excel
Write-Table (Join-Path $outDir 'Insurance_Rollup.xlsx') `
    @('Insurance','Billed','Paid','Unpaid') $insuranceRows @('B','C','D')

# ---- MRN Owes Rollup ----
$excel = Open-Excel
$mrnAgg = @{}

foreach ($f in $files) {
    $wb = $excel.Workbooks.Open($f, $null, $true)
    $ws = $wb.Worksheets.Item(1)
    $used = $ws.UsedRange
    $vals = $used.Value2
    $rows = $used.Rows.Count
    $cols = $used.Columns.Count

    $header = Get-Headers $vals $cols
    $colMrn = $header['Patient MRN']
    $colBal = $header['Current Balance ($)']

    for ($r=2; $r -le $rows; $r++) {
        $mrn = [string]$vals[$r,$colMrn]
        if (-not $mrn) { continue }
        $bal = $vals[$r,$colBal]
        if (-not $bal) { continue }
        if (-not $mrnAgg.ContainsKey($mrn)) { $mrnAgg[$mrn] = 0.0 }
        $mrnAgg[$mrn] += [double]$bal
    }
    $wb.Close($false)
}
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

$mrnRows = $mrnAgg.GetEnumerator() |
    Where-Object { $_.Value -gt 0 } |
    Sort-Object Name |
    ForEach-Object { @($_.Name, [double]$_.Value) }

# Write MRN summary Excel
Write-Table (Join-Path $outDir 'MRN_Owes_Rollup.xlsx') `
    @('MRN','Owes') $mrnRows @('B')

# ---- Monthly Income Rollup ----
$excel = Open-Excel
$monthAgg = @{}

foreach ($f in $files) {
    $wb = $excel.Workbooks.Open($f, $null, $true)
    $ws = $wb.Worksheets.Item(1)
    $used = $ws.UsedRange
    $vals = $used.Value2
    $rows = $used.Rows.Count
    $cols = $used.Columns.Count

    $header = Get-Headers $vals $cols
    $colDate = $header['Posted Date']
    $colAmt = $header['Posted Amount ($)']

    for ($r=2; $r -le $rows; $r++) {
        $dateVal = $vals[$r,$colDate]
        if (-not $dateVal) { continue }

        # Convert Excel serial date or string to DateTime
        $dt = $null
        if ($dateVal -is [double]) {
            $dt = [DateTime]::FromOADate($dateVal)
        } else {
            [DateTime]::TryParse([string]$dateVal, [ref]$dt) | Out-Null
        }
        if (-not $dt) { continue }

        $key = $dt.ToString('yyyy-MM')
        $amt = $vals[$r,$colAmt]
        if (-not $monthAgg.ContainsKey($key)) { $monthAgg[$key] = 0.0 }
        if ($amt) { $monthAgg[$key] += [double]$amt }
    }
    $wb.Close($false)
}
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

$monthRows = $monthAgg.GetEnumerator() |
    Sort-Object Name |
    ForEach-Object { @($_.Name, [double]$_.Value) }

# Write Monthly Income summary Excel
Write-Table (Join-Path $outDir 'Monthly_Income.xlsx') `
    @('Month','Income') $monthRows @('B')

Write-Host "Wrote Insurance_Rollup.xlsx, MRN_Owes_Rollup.xlsx, Monthly_Income.xlsx to $outDir"

# ---- To Run ----
# Open PowerShell and execute:
# powershell -NoProfile -ExecutionPolicy Bypass -File .\annual_rollup.ps1
