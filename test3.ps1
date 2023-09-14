$xlCellTypeLastCell = 2
$startRow = 1
$startCol = 1
$endCol = 3

$excel = New-Object -Com Excel.Application
$wb = $excel.workbooks.open("C:\Users\sumee\Desktop\powershell\sox\excel.xlsx")


$sh = $wb.Sheets.Item(1)
$endRow = $sh.UsedRange.SpecialCells($xlCellTypeLastCell).Row
Write-Host "End Row Value - $($endRow)"
$sheetTitle = "Test"
$rangeAddress = $sh.Cells.Item($startRow, $startCol).Address() + ":" + $sh.Cells.Item($endRow, $endCol).Address()

$data = $sh.Range($rangeAddress).Value2

$data

$newSheet = $wb.Sheets.Add()

# Set the sheet name
$newSheet.Name = $sheetTitle

# Populate the new worksheet with the data
$newSheet.Cells.Item(1, 1).Resize($endRow - $startRow + 1, $endCol).Value2 = $data

$newWorkbookPath = "C:\Users\sumee\Desktop\powershell\sox\new_data3.xlsx"
$wb.SaveAs($newWorkbookPath)

$excel.Workbooks.Close()