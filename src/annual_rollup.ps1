# annual_rollup.ps1
# Generalized version for aggregating Excel data (financial, numeric, or record-based)

function Open-Excel {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    return $excel
}

function Get-Headers($vals, $cols) {
    $header = @{}
    for ($c=1; $c -le $cols; $c++) {
        $name = [string]$vals[1,$c]
        if ($name -and -not $header.ContainsKey($name)) { $header[$name] = $c }
    }
    return $header
}

function Write-Table($outPath, $headers, $rows, $numberCols) {
    $excel = Open-Excel
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)

    # Write headers
    for ($i=0; $i -lt $headers.Count; $i++) {
        $ws.Cells.Item(1, $i+1) = $headers[$i]
    }
    $ws.Range(("A1:" + [char](65 + $headers.Count - 1) + "1")).Font.Bold = $true

    # Write rows
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

    $ws.Columns.AutoFit() | Out-Null
    $wb.SaveAs($outPath)
    $wb.Close($true)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

# Prompt for input files and output directory
$inputFiles = Read-Host "Enter paths to input Excel files (comma-separated)"
$outDir = Read-Host "Enter output folder (e.g., C:\Users\Owner\Downloads)"
$files = $inputFiles -split ',' | ForEach-Object { $_.Trim() }

# ---- Aggregate totals by category ----
$excel = Open-Excel
$categoryAgg = @{}

foreach ($f in $files) {
    $wb = $excel.Workbooks.Open($f, $null, $true)
    $ws = $wb.Worksheets.Item(1)
    $used = $ws.UsedRange
    $vals = $used.Value2
    $rows = $used.Rows.Count
    $cols = $used.Columns.Count

    $header = Get-Headers $vals $cols

    # Example: Generic "Category" and "Amount" columns
    if (-not $header.ContainsKey('Category') -or -not $header.ContainsKey('Amount')) {
        Write-Host "File $f missing required 'Category' or 'Amount' columns. Skipping."
        $wb.Close($false)
        continue
    }

    $colCategory = $header['Category']
    $colAmount = $header['Amount']

    for ($r=2; $r -le $rows; $r++) {
        $cat = [string]$vals[$r,$colCategory]
        if (-not $cat) { $cat = '(Blank)' }
        $amt = $vals[$r,$colAmount]
        if (-not $categoryAgg.ContainsKey($cat)) {
            $categoryAgg[$cat] = 0.0
        }
        if ($amt) { $categoryAgg[$cat] += [double]$amt }
    }

    $wb.Close($false)
}

$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

# Write aggregated output
$rowsOut = $categoryAgg.GetEnumerator() |
    Sort-Object Name |
    ForEach-Object { @($_.Name, [double]$_.Value) }

Write-Table (Join-Path $outDir 'Category_Rollup.xlsx') `
    @('Category','Total') $rowsOut @('B')

Write-Host "Wrote Category_Rollup.xlsx to $outDir"
