$Logfile = "C:\Users\ARMYOFTANKS\path\to\file.xlsx"

#Declare the file path and sheet name
$file = "C:\Users\ARMYOFTANKS\path\to\file.xlsx"
$sheetName = "sheet1"
#Create an instance of Excel.Application and Open Excel file
$objExcel = New-Object -ComObject Excel.Application
$workbook = $objExcel.Workbooks.Open($file)
$sheet = $workbook.Worksheets.Item($sheetName)
$objExcel.Visible=$false
#Count max row
$rowMax = ($sheet.UsedRange.Rows).count
#Declare the starting positions
$rowname,$colname = 1,1
$rowfileURL,$colfileURL = 1,14
$rowfolder,$colfolder = 1,15
#loop to get values and store it
for ($i=1; $i -le $rowMax-1; $i++) {
$name = $sheet.Cells.Item($rowname+$i,$colname).text
$fileURL = $sheet.Cells.Item($rowfileURL+$i,$colfileURL).text
$folder = $sheet.Cells.Item($rowfolder+$i,$colfolder).text
#added a condition to allow skipping of cells and prevent exiting
if ($fileURL -contains 'null') {
    $name >> $Logfile
 }else {
    Invoke-WebRequest $fileURL -OutFile $folder --PreserveFilename
    $folder >> $Logfile
 }

}
#close excel file
$objExcel.quit()
