param(
    [Parameter(Mandatory=$true)][string]$TemplatePath,
    [Parameter(Mandatory=$true)][string]$OutputPath,
    [string]$AsnSheetName = '报关资料与ASN'
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$excel = $null
$book = $null
$sheet = $null
$used = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3
    $snapshots = @()
    foreach ($path in @($TemplatePath, $OutputPath)) {
        $book = $excel.Workbooks.Open($path, 0, $true)
        $snapshot = [ordered]@{}
        foreach ($sheet in $book.Worksheets) {
            if ($sheet.Name -eq $AsnSheetName) { continue }
            $used = $sheet.UsedRange
            # Formula returns constants too, so this compares every input and
            # formula while allowing cached results to follow the new ASN data.
            $snapshot[$sheet.Name] = [ordered]@{
                address = $used.Address()
                contents = @($used.Formula)
                protected = $sheet.ProtectContents
                shapes = $sheet.Shapes.Count
                printArea = $sheet.PageSetup.PrintArea
            }
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($used)
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet)
            $used = $null
            $sheet = $null
        }
        $snapshots += $snapshot
        $book.Close($false)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
        $book = $null
    }
    foreach ($name in $snapshots[0].Keys) {
        $before = ConvertTo-Json -InputObject $snapshots[0][$name] -Depth 10 -Compress
        $after = ConvertTo-Json -InputObject $snapshots[1][$name] -Depth 10 -Compress
        if ($before -cne $after) { throw "Sample sheet changed: $name" }
    }
    Write-Output ('PRESERVED_SHEETS=' + $snapshots[0].Count)
} finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
    foreach ($item in @($used, $sheet, $book, $excel)) {
        if ($null -ne $item -and [Runtime.InteropServices.Marshal]::IsComObject($item)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($item)
        }
    }
}
