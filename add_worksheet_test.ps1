Import-Module -Name ImportExcel

$template = Open-ExcelPackage "scienceTemplate.xlsx"
$wsScience = $template.Workbook.Worksheets["Science"]

$destination = "test.xlsx"

$excel = Open-ExcelPackage $destination

$ws = Add-Worksheet -ExcelPackage $excel -WorkSheetname "Science" -ClearSheet -CopySource $wsScience

Close-ExcelPackage $excel
