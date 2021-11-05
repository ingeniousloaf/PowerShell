# load file
$filepath = "C:\test.xlsx"
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
$wb = $xl.Workbooks.Open($filepath)

# Get data from column 2
$data = $wb.Worksheets['Sheet1'].UsedRange.Rows.Columns[1].Value2

# cleanup
$wb.close()
$xl.Quit()
While([System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) -ge 0){}
while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) -ge 0){} 
Remove-Variable xl,wb # this is optional

# display results
$data | select -skip 1 # remove header
