$xlCellTypeLastCell = 11 
$startRow = 5 
$col = 2 

$excel = New-Object -Com Excel.Application
$wb = $excel.workbooks.open("C:\Users\sumee\Desktop\powershell\sox\excel.xlsx")

for ($i = 1; $i -le $wb.Sheets.Count; $i++)
{
    $sh = $wb.Sheets.Item($i)
    $endRow = $sh.UsedRange.SpecialCells($xlCellTypeLastCell).Row
    # $city = $sh.Cells.Item($startRow, $col).Value2
    $city = $wb.Sheets.Item($i).name
    $rangeAddress = $sh.Cells.Item($startRow, $col).Address() + ":" + $sh.Cells.Item($endRow, $col).Address()
    $sh.Range($rangeAddress).Value2 | foreach {
        New-Object PSObject -Property @{City = $city; Area = $_ }
    }
}

$excel.Workbooks.Close()