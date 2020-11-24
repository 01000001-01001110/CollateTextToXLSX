#Create An Excel File
$excel = New-Object -ComObject excel.application
$excel.visible = $True

#Add Workbook
$workbook = $excel.Workbooks.Add()

<#Rename Workbook
$workbook= $workbook.Worksheets.Item(1)
$workbook.Name = 'Client name and #'#>

#create the column headers
$workbook.Cells.Item(1,1) = 'Client name and n°'
$workbook.Cells.Item(1,2) = 'OK'
$workbook.Cells.Item(1,3) = 'Disconnected'
$workbook.Cells.Item(1,4) = 'Unresponsive'
$workbook.Cells.Item(1,5) = 'Unreachable'
$workbook.Cells.Item(1,6) = 'Version'
$workbook.Cells.Item(1,7) = 'Date Gathered'

$move = "C:\Users\iNet\Desktop\Testing"
$root = "C:\Users\iNet\Desktop\Testing"
$files = Get-ChildItem -Path $root -Filter *.txt

#Starting on Row 2
[int]$i = 2
ForEach ($file in $files){
   
$location = $root+"\"+$file

#Format your client data to output what you want to see. 
$ClientData = select-string -path "$location" -pattern "Client"
$ClientData = $ClientData.line
$ClientData = $ClientData -replace "Client n° :" -replace ""
$ClientData = $ClientData -replace "Client name :" -replace "|"
$row = $i
$Column = 1
$workbook.Cells.Item($row,$column)= "$ClientData"

#Data Read Date
$DataReadDate = select-string -path "$location" -pattern "Data read"
$DataReadDate = $DataReadDate.line
$DataReadDate = $DataReadDate -replace "Data read " -replace ""
    #Data Read Date, you asked for everything but this.
$row = $i
$Column = 7
$workbook.Cells.Item($row,$column)= "$DataReadDate" 

#Version
$Version = select-string -path "$location" -pattern "Version:"
$Version = $Version.line
$Version = $Version -replace "Version: " -replace ""
$row = $i
$Column = 6
$workbook.Cells.Item($row,$column)= "$Version"

#How Many Times Unresponsive Shows Up
$Unresponsive = (Get-Content "$location" | select-string -pattern "Unresponsive").length
$row = $i
$Column = 4
$workbook.Cells.Item($row,$column)= "$Unresponsive"
 
#How Many Times Disconnected Shows Up
$Disconnected = (Get-Content "$location" | select-string -pattern "Disconnected").length
$row = $i
$Column = 3
$workbook.Cells.Item($row,$column)= "$Disconnected"

#How Many Times Unreachable host Shows Up
$Unreachable = (Get-Content "$location" | select-string -pattern "Unreachable host").length
$row = $i
$Column = 5
$workbook.Cells.Item($row,$column)= "$Unreachable"

#How Many Times OK Shows Up
$OK = (Get-Content "$location" | select-string -pattern "OK").length
$row = $i
$Column = 2
$workbook.Cells.Item($row,$column)= "$OK"
#Iterate by one so each text file goes to its own line.
$i++
}

#Save Document
$output = "\Output.xlsx"
$FinalOutput = $move+$output
#saving & closing the file
$workbook.SaveAs($move)
$excel.Quit()
