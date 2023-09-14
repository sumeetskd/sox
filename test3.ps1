$xlCellTypeLastCell = 11 
$startRow = 1
$col = 2

$excel = New-Object -Com Excel.Application
$wb = $excel.workbooks.open("C:\Users\sumee\Desktop\powershell\sox\new_data.xlsx")

# for ($i = 1; $i -le $wb.Sheets.Count; $i++)
# {
#     $sh = $wb.Sheets.Item($i)
#     $endRow = $sh.UsedRange.SpecialCells($xlCellTypeLastCell).Row
#     $city = $sh.Cells.Item($startRow, $col).Value2
#     $rangeAddress = $sh.Cells.Item($startRow + 1, $col).Address() + ":" + $sh.Cells.Item($endRow, $col).Address()
#     $sh.Range($rangeAddress).Value2 | foreach {
#         New-Object PSObject -Property @{City = $city; Area = $_ }
#     }
# }

$sh = $wb.Sheets.Item(1)
$endRow = $sh.UsedRange.SpecialCells($xlCellTypeLastCell).Row
$city = $sh.Cells.Item($startRow, $col).Value2
$rangeAddress = $sh.Cells.Item($startRow, $col).Address() + ":" + $sh.Cells.Item($endRow, $col).Address()
# $sh.Range($rangeAddress).Value2 | foreach {
#     New-Object PSObject -Property @{City = $city; Area = $_ }
# }
$data = $sh.Range($rangeAddress).Value2

$data

# $newSheet = $wb.Sheets.Add()

# # Set the sheet name to the city
# $newSheet.Name = $city

# # Populate the new worksheet with the data
# $newSheet.Cells.Item(1, 1).Resize($endRow - $startRow + 1, 1).Value2 = $data

# $newWorkbookPath = "C:\Users\sumee\Desktop\powershell\sox\new_data.xlsx"
# $wb.SaveAs($newWorkbookPath)

$excel.Workbooks.Close()