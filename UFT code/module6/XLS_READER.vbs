
option explicit 
Dim objXls, myXls, sheet, exitingSheeting,fso
Const ORFilePath=environment("projectPath")&"/OR/OR.xls"
Const ORSheet="Objects"

'  Read data from the xls
Function readData(filePath, sheetName, row,col)
   Set objXls = createobject("Excel.application")  '' microsoft  xl 
   objXls.Visible= false
   objXls.DisplayAlerts=false
   ' exit the function if file is not present
	If NOT isFilePresent(filePath) Then
		msgbox "FILE NOT PRESENT " & filePath
		Exit function
	End If

	Set myXls = objXls.Workbooks.Open (filePath)
	Set sheet = myXls.Sheets(sheetName)
	readData = sheet.cells(row,col).value   
	' close everything
    myXls.Close
	objXls.Quit

' destroy object
  Set objXls=nothing
  Set myXls=nothing
  Set sheet=nothing
End Function

'Write data in the xls file
Function writeData(filePath,sheetName,row,col,data)
   Set objXls = createobject("Excel.application")  '' microsoft  xl 
   objXls.Visible= false
   objXls.DisplayAlerts=false
   ' exit the function if file is not present
	If NOT isFilePresent(filePath) Then
		msgbox "FILE NOT PRESENT " & filePath
		Exit function
	End If

	Set myXls = objXls.Workbooks.Open (filePath)
	Set sheet = myXls.Sheets(sheetName)
	sheet.cells(row,col).delete
	sheet.cells(row,col).value = data
	' close everything
	myXls.Save
    myXls.Close
	objXls.Quit
	 'destroy object
  Set objXls=nothing
  Set myXls=nothing
  Set sheet=nothing
End Function


' to add link in the xls file
Function writeLink(filePath,sheetName,row,col,linkText,location)
   print "writing data"
   Set objXls = createobject("Excel.application")  '' microsoft  xl 
   objXls.Visible= false
   objXls.DisplayAlerts=false
   ' exit the function if file is not present
	If NOT isFilePresent(filePath) Then
		msgbox "FILE NOT PRESENT " & filePath
		Exit function
	End If

	Set myXls = objXls.Workbooks.Open (filePath)
	Set sheet = myXls.Sheets(sheetName)
	sheet.cells(row,col).delete
    sheet.cells(row,col).value  = "=Hyperlink("&chr(34)&location&chr(34)&", "&chr(34)&linkText&chr(34)&")"
	' close everything
	myXls.Save
    myXls.Close
	objXls.Quit
	 'destroy object
  Set objXls=nothing
  Set myXls=nothing
  Set sheet=nothing
End Function

'  Returns total rows in a sheet
Function getRowCount(filePath,sheetName)
  

'  all the cells for all the cols are empty
' know total cols in the sheet
Dim cols,isRowPresent,rows,isLastRow
cols = getColumnCount(filePath,sheetName)
rows=1 '  assuming there is one row
isRowPresent = true
isLastRow= false

While isRowPresent
' check to see if all the cells are empty or not
Dim colNum
	For colNum=1 to cols
		If  readData(filePath,sheetName,rows,colNum) <> ""  Then
			' this is not the last row
			rows=rows+1
			Exit for
		End If
'       if you have just read the last cell
			If  colNum = cols Then
				' this is the last row
				isLastRow = true
				Exit for
			End If

	Next

	' check if  its the last row
	If  isLastRow Then
		isRowPresent=false
	End If
Wend

getRowCount=  (rows-2)


End Function



' return total cols in a sheet
Function getColumnCount(filePath,sheetName)
Dim colCount
colCount=0 '  assume no cols
While trim(readData(filePath,sheetName,1,(colCount+1))) <> ""
colCount = colCount+1
Wend
getColumnCount= colCount

End Function



' function returns true if  file is present and false if not present
Function isFilePresent(filePath)
Set fso = createobject("Scripting.FileSystemObject")
If  fso.FileExists(filePath)  Then
		isFilePresent = true
else
       isFilePresent=false
End If
' destroy object
Set fso= nothing
End Function






''''''''''''''''''''''''''''''''PRE''''''''''''''''''''''''''''''''''
'  if sheet is existing
Public function isSheetExit (xlsFIilePath,sheetName)
Set  objExcel = CreateObject("Excel.Application")
Set  objWorkbook = objExcel.Workbooks.Open(xlsFIilePath)


totalSheets =  objWorkbook.Worksheets.count
For i=1 to totalSheets
	If  objWorkbook.Worksheets(i).name = sheetName Then
		isSheetExit=true

			objWorkbook.Close
			objExcel.Application.Quit
			Set  objWorkbook=nothing
			Set  objExcel=nothing

		Exit function
	End If
Next

isSheetExit=False
'objWorkbook.Save
objWorkbook.Close
objExcel.Application.Quit

Set  objWorkbook=nothing
Set  objExcel=nothing
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' if col is existing or not
Public function iscolumnExist (xlsFilePath,sheetName,colName)
   print xlsFIilePath
   Set  objExcel = CreateObject("Excel.Application")
   Set  objWorkbook = objExcel.Workbooks.Open(xlsFIilePath)

	totalCols=getColumnCount(xlsFilePath,sheetName)

	For colNum=1 to  totalCols
		cuurentColName = objSheet.Cells(1,colNum)

		If cuurentColName =  colName Then
			iscolumnExist=True

			objWorkbook.Close
			objExcel.Application.Quit
			Set  objWorkbook=nothing
			Set  objExcel=nothing

			Exit function
		End If

	Next

iscolumnExist = false
objWorkbook.Close
objExcel.Application.Quit

Set  objWorkbook=nothing
Set  objExcel=nothing

End Function

'' 'Delete a col and all values in that column

