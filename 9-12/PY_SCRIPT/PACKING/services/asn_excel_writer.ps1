param([Parameter(Mandatory=$true)][string]$PlanPath)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$plan = Get-Content -LiteralPath $PlanPath -Raw -Encoding UTF8 | ConvertFrom-Json
$excel = $null
$book = $null
$asnExcelPid = 0
$existingExcelPids = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class AsnExcelProcess {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
}
'@
try {
    $excel = New-Object -ComObject Excel.Application
    [uint32]$createdPid = 0
    [void][AsnExcelProcess]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$createdPid)
    if ($existingExcelPids -notcontains [int]$createdPid) { $asnExcelPid = [int]$createdPid }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open([string]$plan.template, 0, $true)
    $sheet = $book.Worksheets.Item([string]$plan.sheet)
    $protected = $sheet.ProtectContents
    if ($protected -and -not $plan.data_last_row) { throw 'ASN工作表受保护，无法填写。' }
    # Keep every sample sheet, including its formulas, formatting and tab order.
    if ($plan.scope -ne 'all') { throw 'ASN输出必须保留模板中的全部工作表。' }
    $count = $plan.rows.Count
    if ($count -lt 1) { throw '没有ASN明细。' }
    $lastColumn = [string]$plan.last_column
    $clearLastColumn = [string]$plan.clear_last_column
    $columnCount = [int]$plan.columns
    if ($columnCount -lt 1) { throw 'ASN模板列配置无效。' }
    $lastRow = 25 + $count
    if ($plan.data_last_row -and $lastRow -gt [int]$plan.data_last_row) {
        throw 'ASN明细超过模板开放填写的行数。'
    }
    $templateLastRow = $sheet.UsedRange.Row + $sheet.UsedRange.Rows.Count - 1
    $clearLast = [Math]::Max($lastRow, $templateLastRow)
    if ($plan.data_last_row) { $clearLast = [int]$plan.data_last_row }
    if ($protected -and $sheet.Range('A26:' + $clearLastColumn + $clearLast).Locked -ne $false) {
        throw 'ASN明细区域包含锁定单元格，未修改工作表保护。'
    }
    # Keep all existing row-specific formats. Only extend beyond the template's range.
    if ($lastRow -gt $templateLastRow) {
        $firstNewRow = [Math]::Max(27, $templateLastRow + 1)
        $sheet.Range('A27:' + $lastColumn + '27').Copy()
        $sheet.Range('A' + $firstNewRow + ':' + $lastColumn + $lastRow).PasteSpecial(-4122)
        $sheet.Range('A' + $firstNewRow + ':' + $lastColumn + $lastRow).RowHeight = $sheet.Rows.Item(27).RowHeight
    }
    $sheet.Range('A26:' + $clearLastColumn + $clearLast).ClearContents()
    foreach ($address in $plan.clear_cells) { $sheet.Range([string]$address).ClearContents() }
    $matrix = New-Object 'object[,]' $count,$columnCount
    for ($r = 0; $r -lt $count; $r++) {
        for ($c = 0; $c -lt $columnCount; $c++) {
            $value = $plan.rows[$r][$c]
            if ($null -eq $value) { $matrix[$r,$c] = $null }
            elseif ($value -is [string]) { $matrix[$r,$c] = "'" + $value }
            else { $matrix[$r,$c] = [double]$value }
        }
    }
    $body = $sheet.Range('A26:' + $lastColumn + $lastRow)
    # Write the array before mixed scalar header assignments (PowerShell COM binding).
    $body.Value2 = $matrix
    foreach ($property in $plan.headers.PSObject.Properties) {
        $target = $sheet.Range($property.Name)
        if ($target.MergeCells) { $target = $target.MergeArea.Cells.Item(1,1) }
        if ($target.MergeCells) { $target.MergeArea.ClearContents() } else { $target.ClearContents() }
        if ($null -ne $property.Value -and [string]$property.Value -ne '') {
            if ($plan.date_cells -contains $property.Name) {
                $date = [DateTime]::ParseExact([string]$property.Value, 'yyyy-MM-dd',
                                             [Globalization.CultureInfo]::InvariantCulture)
                $serial = $date.ToOADate()
                if ($book.Date1904) { $serial -= 1462 }
                if ([string]$target.NumberFormat -notmatch '(?i)[ymd]') {
                    $target.NumberFormat = 'yyyy-mm-dd'
                }
                $target.Value2 = [double]$serial
            } elseif ($property.Value -is [string]) {
                # Preserve identifiers/leading zeroes without changing the template format.
                $target.Value2 = "'" + [string]$property.Value
            } else {
                if ($target.NumberFormat -eq '@') { $target.NumberFormat = 'General' }
                $target.Value2 = [double]$property.Value
            }
        }
    }
    foreach ($formula in $plan.formulas) {
        $sheet.Range([string]$formula.cell).Formula = [string]$formula.value
    }
    # Preserve an unset or already sufficient print area; extend a defined area only as needed.
    if ($plan.extend_print_area -and $sheet.PageSetup.PrintArea) {
        $printRange = $sheet.Range($sheet.PageSetup.PrintArea)
        if ($printRange.Areas.Count -eq 1 -and $lastRow -gt ($printRange.Row + $printRange.Rows.Count - 1)) {
            $sheet.PageSetup.PrintArea = $printRange.Resize($lastRow - $printRange.Row + 1,
                                                          $printRange.Columns.Count).Address()
        }
    }
    $sheet.Activate()
    $sheet.Range('A1').Select()
    $excel.CalculateFullRebuild()
    # Validate the ASN we fill. Other sheets retain the sample's existing formulas,
    # including pre-existing errors; do not repair them or block this export.
    $errors = New-Object System.Collections.Generic.List[string]
    $used = $sheet.Range('A1:' + $lastColumn + $lastRow)
    # Read error codes in one COM call; inspect text only for possible errors.
    $values = $used.Value2
    for ($r = 1; $r -le $values.GetLength(0); $r++) {
        for ($c = 1; $c -le $values.GetLength(1); $c++) {
            $value = $values[$r,$c]
            if ($value -is [System.Runtime.InteropServices.ErrorWrapper] -or
                ($value -is [int] -and $value -lt 0)) {
                $cell = $used.Cells.Item($r,$c)
                $errors.Add($sheet.Name + '!' + $cell.Address($false,$false) + '=' + $cell.Text)
                if ($errors.Count -ge 10) { break }
            }
        }
        if ($errors.Count -ge 10) { break }
    }
    if ($errors.Count -gt 0) { throw ('ASN页存在失效公式，未生成文件：' + ($errors -join '；')) }
    $book.CheckCompatibility = $false
    if ([int]$plan.file_format -eq 52 -and -not $book.HasVBProject) {
        throw '星辉仓模板缺少宏项目，未生成文件。'
    }
    $book.SaveAs([string]$plan.destination, [int]$plan.file_format)
    Write-Output 'ASN_SAVED'
} catch {
    if ($null -eq $excel) { [Console]::Error.WriteLine('生成 ASN 需要本机安装可用的 Microsoft Excel。') }
    else { [Console]::Error.WriteLine($_.Exception.Message) }
    exit 1
} finally {
    try {
        if ($null -ne $book) { $book.Close($false) }
    } finally {
        if ($null -ne $excel) { $excel.Quit() }
        foreach ($comObject in @($body, $printRange, $target, $used, $cell, $sheet, $book, $excel)) {
            if ($null -ne $comObject -and [Runtime.InteropServices.Marshal]::IsComObject($comObject)) {
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject)
            }
        }
        # Excel COM can retain hidden helper processes after Quit. Only clean up the
        # newly created instance, never an Excel process that predated this export.
        if ($asnExcelPid -gt 0) {
            $ownedProcess = Get-Process -Id $asnExcelPid -ErrorAction SilentlyContinue
            if ($null -ne $ownedProcess -and -not $ownedProcess.WaitForExit(3000)) {
                Stop-Process -Id $asnExcelPid -ErrorAction SilentlyContinue
            }
        }
    }
}
