param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,
    [Parameter(Mandatory = $true)]
    [string]$CaseName
)

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $workbook = $excel.Workbooks.Open((Resolve-Path -LiteralPath $WorkbookPath).Path, 0, $false)

    function Set-PrintArea($sheet, [string]$area, [int]$orientation, [int]$paperSize) {
        $sheet.PageSetup.PrintArea = $area
        $sheet.PageSetup.Orientation = $orientation
        $sheet.PageSetup.PaperSize = $paperSize
        $sheet.PageSetup.Zoom = $false
        $sheet.PageSetup.FitToPagesWide = 1
        $sheet.PageSetup.FitToPagesTall = 1
        $sheet.PageSetup.LeftMargin = 36
        $sheet.PageSetup.RightMargin = 36
        $sheet.PageSetup.TopMargin = 36
        $sheet.PageSetup.BottomMargin = 36
    }

    switch ($CaseName) {
        "excel_printareas_shapes" {
            $sheet = $workbook.Worksheets.Item("PrintAreas")
            $sheet.Rows.Item(10).Hidden = $true
            $sheet.Columns.Item("G").Hidden = $true
            Set-PrintArea $sheet '$B$3:$F$6,$H$20:$J$21' 2 3
        }
        "excel_chunks" {
            Set-PrintArea $workbook.Worksheets.Item("Chunks") '$A$1:$AN$121' 2 3
        }
        "excel_max_width" {
            Set-PrintArea $workbook.Worksheets.Item("MaxWidth") '$A$1:$BZT$9' 2 26
        }
        "excel_oversized_error" {
            Set-PrintArea $workbook.Worksheets.Item("Normal") '$A$1:$L$6' 2 3
            Set-PrintArea $workbook.Worksheets.Item("Oversized") '$A$1:$SR$4' 2 26
            Set-PrintArea $workbook.Worksheets.Item("MaxColumns") '$A$1:$XFD$2' 2 26
        }
        "excel_oversized_skip" {
            Set-PrintArea $workbook.Worksheets.Item("KeepMe") '$A$1:$L$11' 2 3
            Set-PrintArea $workbook.Worksheets.Item("SkipMe") '$A$1:$SR$4' 2 26
        }
        "excel_empty_shapes" {
            Set-PrintArea $workbook.Worksheets.Item("ShapesOnly") '$A$1:$L$18' 2 3
            $workbook.Worksheets.Item("HiddenCandidate").Visible = 0
        }
        "excel_empty_workbook" {
            $workbook.Worksheets.Item("HiddenCandidate").Visible = 0
        }
        "excel_orientation_margin" {
            Set-PrintArea $workbook.Worksheets.Item("Portrait") '$A$1:$H$26' 1 9
            Set-PrintArea $workbook.Worksheets.Item("Landscape") '$A$1:$R$26' 2 3
        }
        "excel_quality_pagination" {
            $qualitySheets = @(
                @("WideOnly", '$A$1:$BL$13'),
                @("TallOnly", '$A$1:$H$181'),
                @("WideTall", '$A$1:$AV$81')
            )
            foreach ($qualitySheet in $qualitySheets) {
                $sheet = $workbook.Worksheets.Item($qualitySheet[0])
                Set-PrintArea $sheet $qualitySheet[1] 1 9
                $sheet.PageSetup.PrintTitleRows = '$1:$1'
                $sheet.PageSetup.PrintTitleColumns = '$A:$A'
            }
        }
        "excel_paper_forms" {
            $forms = @(
                @("A4Portrait", 1, 9), @("LetterPortrait", 1, 1), @("B4Portrait", 1, 12),
                @("TabloidPortrait", 1, 3), @("A3Portrait", 1, 8), @("LetterLandscape", 2, 1),
                @("LegalLandscape", 2, 5), @("LedgerLandscape", 2, 4),
                @("ArchCLandscape", 2, 24), @("ArchDLandscape", 2, 25), @("ArchELandscape", 2, 26)
            )
            foreach ($form in $forms) {
                Set-PrintArea $workbook.Worksheets.Item($form[0]) '$A$1:$B$5' $form[1] $form[2]
            }
        }
        default { throw "Unknown fixture case: $CaseName" }
    }
    $excel.PrintCommunication = $true
    $workbook.Save()
} finally {
    if ($workbook) { $workbook.Close($true) }
    if ($excel) { $excel.Quit() }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
