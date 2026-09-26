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
    if ($sheet.ProtectContents) { throw 'ASN工作表受保护，无法填写。' }
    if ($plan.scope -eq 'asn_only') {
        for ($i = $book.Worksheets.Count; $i -ge 1; $i--) {
            if ($book.Worksheets.Item($i).Name -ne $plan.sheet) { $book.Worksheets.Item($i).Delete() }
        }
    } elseif ($plan.scope -ne 'all') { throw '未确认输出工作表范围。' }
    $count = $plan.rows.Count
    if ($count -lt 1) { throw '没有ASN明细。' }
    $lastRow = 25 + $count
    $clearLast = [Math]::Max($lastRow, $sheet.UsedRange.Row + $sheet.UsedRange.Rows.Count - 1)
    # Keep the first-line highlights; extend the normal body style for all subsequent lines.
    if ($count -gt 1) {
        $sheet.Range('A27:AU27').Copy()
        $sheet.Range('A27:AU' + $lastRow).PasteSpecial(-4122)
        $sheet.Range('A27:AU' + $lastRow).RowHeight = $sheet.Rows.Item(27).RowHeight
    }
    $sheet.Range('A26:AV' + $clearLast).ClearContents()
    $sheet.Range('I24').ClearContents()
    foreach ($property in $plan.headers.PSObject.Properties) {
        $target = $sheet.Range($property.Name)
        if ($target.MergeCells) { $target = $target.MergeArea.Cells.Item(1,1) }
        if ($target.MergeCells) { $target.MergeArea.ClearContents() } else { $target.ClearContents() }
        if ($null -ne $property.Value -and [string]$property.Value -ne '') {
            $target.NumberFormat = '@'
            $target.Value2 = [string]$property.Value
        }
    }
    $matrix = New-Object 'object[,]' $count,47
    for ($r = 0; $r -lt $count; $r++) {
        for ($c = 0; $c -lt 47; $c++) {
            $value = $plan.rows[$r][$c]
            if ($null -eq $value) { $matrix[$r,$c] = $null }
            elseif ($value -is [string]) { $matrix[$r,$c] = "'" + $value }
            else { $matrix[$r,$c] = [double]$value }
        }
    }
    $sheet.Range('A26:AU' + $lastRow).Value2 = $matrix
    foreach ($formula in $plan.formulas) {
        $sheet.Range([string]$formula.cell).Formula = [string]$formula.value
    }
    $sheet.PageSetup.PrintArea = 'A1:AU' + $lastRow
    $sheet.Activate()
    $sheet.Range('A1').Select()
    $excel.CalculateFullRebuild()
    # Validate displayed Excel errors without relying on SpecialCells on protected sheets.
    $errors = New-Object System.Collections.Generic.List[string]
    foreach ($checkSheet in $book.Worksheets) {
        if ($plan.scope -eq 'asn_only' -or $checkSheet.Name -ne 'ASN填写指引') {
            $used = $checkSheet.UsedRange
            if ($checkSheet.Name -eq $plan.sheet) { $used = $checkSheet.Range('A1:AU' + $lastRow) }
            # Read error codes in one COM call; inspect text only for possible errors.
            $values = $used.Value2
            for ($r = 1; $r -le $values.GetLength(0); $r++) {
                for ($c = 1; $c -le $values.GetLength(1); $c++) {
                    $value = $values[$r,$c]
                    if ($value -is [System.Runtime.InteropServices.ErrorWrapper] -or
                        ($value -is [int] -and $value -lt 0)) {
                        $cell = $used.Cells.Item($r,$c)
                        $errors.Add($checkSheet.Name + '!' + $cell.Address($false,$false) + '=' + $cell.Text)
                        if ($errors.Count -ge 10) { break }
                    }
                }
                if ($errors.Count -ge 10) { break }
            }
        }
        if ($errors.Count -ge 10) { break }
    }
    if ($errors.Count -gt 0) { throw ('模板存在失效公式，未生成文件：' + ($errors -join '；')) }
    $book.CheckCompatibility = $false
    $book.SaveAs([string]$plan.destination, 56)
    Write-Output 'ASN_SAVED'
} catch {
    if ($null -eq $excel) { [Console]::Error.WriteLine('生成 .xls 需要本机安装可用的 Microsoft Excel。') }
    else { [Console]::Error.WriteLine($_.Exception.Message) }
    exit 1
} finally {
    try {
        if ($null -ne $book) { $book.Close($false) }
    } finally {
        if ($null -ne $excel) { $excel.Quit() }
        foreach ($comObject in @($target, $used, $checkSheet, $cell, $sheet, $book, $excel)) {
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
