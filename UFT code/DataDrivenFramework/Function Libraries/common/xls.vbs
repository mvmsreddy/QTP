Option explicit
Dim xlApp,myXLS,sheet,sheetName,xlPath,fso
Set xlApp= CreateObject("Excel.Application")
   xlApp.Visible=false
   xlApp.DisplayAlerts=false
Set sheet= nothing
Set myXLS= nothing


'  saves the xl file and destroys all objectects related to it
Function destroyFile
   ' destroys sheet
If NOT sheet is nothing Then
	Set sheet=nothing
End If
' destroy xl file
If NOT  myXLS is nothing Then
 myXLS.save
 myxls.close
 Set myxls=nothing
End If
End Function


Function destroyXLSApp
destroyFile
If  NOT xlApp is nothing Then
	xlApp.Application.Quit
	Set xlApp=nothing
End If
End Function

'  Reads the data from XLS File
Function readData(xlFilePath,sName,row,col)

	If  xlPath <> xlFilePath Then
		' destroy previous xls opened - if
			destroyFile
		'  check if the file is existing
			If NOT isFileExisiting(xlFilePath) Then
				msgbox "File not found " & xlFilePath
				exitTest
			End If
		' open the xl file
		Set myXls = xlApp.Workbooks.Open(xlFilePath)
		xlPath=xlFilePath

		' check if sheet is present
		If NOT isSheetExisting(xlFilePath,sName) Then
				msgbox  xlFilePath & " has not got sheet -  " & sName
				exitTest
		End If
		' open the sheet of xl file
		Set sheet=myXls.sheets(sName)
		sheetName=sName

	'  file is same but sheet is diff
    ElseIf sheetName <>  sName Then
	   ' check if sheet is present
		If NOT isSheetExisting(xlFilePath,sName) Then
				msgbox  xlFilePath & " has not got sheet -  " & sName
				exitTest
		End If

		   ' destroys sheet
				If NOT sheet is nothing Then
			Set sheet=nothing
			End If

		' open the sheet of xl file
		Set sheet=myXls.sheets(sName)
		sheetName=sName

	End If

' read the data from the sheet

  readData = sheet.cells(row,col)
End Function


' write data in xls file
Function writeData(xlFilePath,sName,row,col,data)
If  xlPath <> xlFilePath Then
		' destroy previous xls opened - if
			destroyFile
		'  check if the file is existing
			If NOT isFileExisiting(xlFilePath) Then
				msgbox "File not found " & xlFilePath
				exitTest
			End If
		' open the xl file
		Set myXls = xlApp.Workbooks.Open(xlFilePath)
		xlPath=xlFilePath

		' check if sheet is present
		If NOT isSheetExisting(xlFilePath,sName) Then
				msgbox  xlFilePath & " has not got sheet -  " & sName
				exitTest
		End If
		' open the sheet of xl file
		Set sheet=myXls.sheets(sName)
		sheetName=sName

	'  file is same but sheet is diff
    ElseIf sheetName <>  sName Then
	   ' check if sheet is present
		If NOT isSheetExisting(xlFilePath,sName) Then
				msgbox  xlFilePath & " has not got sheet -  " & sName
				exitTest
		End If

		   ' destroys sheet
				If NOT sheet is nothing Then
			Set sheet=nothing
			End If

		' open the sheet of xl file
		Set sheet=myXls.sheets(sName)
		sheetName=sName

	End If
' write data
sheet.cells(row,col).value=data

End Function

'returns total cols in xls file
Function getColumnCount(xlFilePath,sheetName)
Dim totalCols
totalCols=0
While readData(xlFilePath,sheetName,1,(totalCols+1)) <> ""
	totalCols=totalCols+1
Wend
getColumnCount=totalCols

End Function

' Returns true if col is existing
Function IsColumnExisting(xlFilePath,sheetName,cName)
Dim cId,colName
For cId=1 to getColumnCount(xlFilePath,sheetName)
	colName = readData(xlFilePath,sheetName,1,cId) 
		If  colName = cName Then
			IsColumnExisting=true
			Exit function
		End If
Next
			IsColumnExisting=false
End Function

'Returns total Rows
Function getRowCount(xlFilePath,sheetName)
   Dim totalRows,cId
totalRows=0
For cId=1 to getColumnCount(xlFilePath,sheetName)

  While readData(xlFilePath,sheetName,(totalRows+1),cId) <> ""
		totalRows=totalRows+1
  Wend

Next
getRowCount= totalRows

End Function

' Add a sheet to the xls file
Function addSheet(xlFilePath,sName)

   If  xlPath <> xlFilePath Then
		' destroy previous xls opened - if
			destroyFile
		'  check if the file is existing
			If NOT isFileExisiting(xlFilePath) Then
				msgbox "File not found " & xlFilePath
				exitTest
			End If
		' open the xl file
		Set myXls = xlApp.Workbooks.Open(xlFilePath)
		xlPath=xlFilePath

	If isSheetExisting(xlFilePath,sName) Then
		msgbox "Sheet already exists "& sheetName
		exittest
	End If
end if
	
	Set sheet=myXls.sheets.add
	sheet.name=sName
	sheetName=sName


End Function

' checks if a file exists
Function isFileExisiting(filePath)
   Set fso = createObject("Scripting.FileSystemObject")

   If  fso.FileExists(filePath)  Then
	isFileExisiting=true
   else
	isFileExisiting=false
   End If

End Function

'  Checks if sheet is existing
Function isSheetExisting(filePath,sName)
   Dim totalsheets,sNum
	totalsheets = myXLS.Worksheets.count
	For sNum=1 to totalsheets
		If  myXLS.Worksheets(sNum).name = sName Then
			isSheetExisting=true
			Exit Function
		End If
	Next
isSheetExisting=false
End Function

' compare the sheets
Function compareSheets(xlFilePath,s1,s2,resultSheetName)
' max row,col
Dim rows,cols,rNum,cNum

If getRowCount(xlFilePath,s1) >=  getRowCount(xlFilePath,s2)  Then
	rows=getRowCount(xlFilePath,s1) 
else
    rows=getRowCount(xlFilePath,s2)
End If

If getColumnCount(xlFilePath,s1) >= getColumnCount(xlFilePath,s2) Then
	cols=getColumnCount(xlFilePath,s1)
else
    cols=getColumnCount(xlFilePath,s2)
End If

addSheet xlFilePath,resultSheetName

For rNum=1 to rows

   For cNum=1 to cols

		If readData(xlFilePath,s1,rNum,cNum) = readData(xlFilePath,s2,rNum,cNum)  Then
			writeData xlFilePath,resultSheetName,rNum,cNum,"Y"
		else
			writeData xlFilePath,resultSheetName,rNum,cNum,"N"
		End If

   Next

Next

End Function

Function getData(xlFilePath,sName,testCaseName)
   Dim testStartIndex,totalRows,colStartIndex,cols,rowStartIndex,rows,r,c,cellData,data,rData,dIndex

testStartIndex=0
' find the location of the test
Do 
		If testCaseName = readData(xlFilePath,sName,(testStartIndex+1),1) Then
			testStartIndex=testStartIndex+1
			Exit do
		else
			testStartIndex=testStartIndex+1
		End If
Loop While true


'msgbox  "Test starts at " & testStartIndex
' find cols
colStartIndex=testStartIndex+1
cols=0
While readData(xlFilePath,sName,colStartIndex,(cols+1)) <> ""
cols=cols+1
Wend
'msgbox "Total Cols for the test are " & cols
' find rows
rowStartIndex=testStartIndex+2
rows=0
While readData(xlFilePath,sName,(rowStartIndex),1) <> ""
rows=rows+1
rowStartIndex=rowStartIndex+1
wend
If  rows = 0 Then
rows=rows+1
End If
rowStartIndex=testStartIndex+2
'msgbox  "Total Test Rows are "& rows
dIndex=1
'print rowStartIndex
Set data=createObject("Scripting.Dictionary")
For r=rowStartIndex to (rowStartIndex+rows-1)
'msgbox "here"
Set rData=createObject("Scripting.Dictionary")
	For c=1 to cols
		cellData=cellData& "  "&  readData(xlFilePath,sName,r,c)
		rData.Add readData(xlFilePath,sName,colStartIndex,c),readData(xlFilePath,sName,r,c)
	Next
'cellData=cellData&vbcrlf
data.Add dIndex,rData
dIndex=dIndex+1
Next
'msgbox cellData
Set getData=data
End Function


' Logging
Function LogThis(msg)
   'print "xx " & environment("logFileName") &"  "& msg
   If logFilePath="" Then
	   logFilePath=environment("logFilePath")
   End If
   Dim fso,file
   Const ForAppending=8
   Set fso=createObject("Scripting.FileSystemObject")
  Set file=fso.OpenTextFile(logFilePath,ForAppending,true)
  file.WriteLine now&" -- " & msg
  file.Close
  Set file=nothing
  Set fso=nothing
End Function



'''''''''''''''''''Reporting error in xls and taking screenshot'''''''''''''''''''
Function reportError(testName,testRowNum,errMsg)
   Dim dataRows,testStartRowNum,resultColNum,screenshotFileName,rNum,currentValue
	screenshotFileName = environment("Project_Path")&"screenshots\"&testName&"_"&testRowNum&".png"
	desktop.CaptureBitmap screenshotFileName,1
   ' first find the row from which test starts
		rNum=1
		'Be careful
		while readData(environment("ResultxlsPath"),environment("testDataSheet"),rNum,1) <>  testName 
		rNum=rNum+1
		wend
	testStartRowNum = rNum
	resultColNum=1
	While readData(environment("ResultxlsPath"),environment("testDataSheet"),testStartRowNum+1,resultColNum) <> "Result"
		resultColNum =resultColNum+1
	Wend

	currentValue = readData(environment("ResultxlsPath"),environment("testDataSheet"),testStartRowNum+1+testRowNum,resultColNum)
	writeData environment("ResultxlsPath"),environment("testDataSheet"),testStartRowNum+1+testRowNum,resultColNum,currentValue & vbnewline & "------" & vbnewline & errMsg
	writeData environment("ResultxlsPath"),environment("testDataSheet"),testStartRowNum+1+testRowNum,resultColNum+1, "=hyperlink("&chr(34)&"file:///"&screenshotFileName&chr(34)&","&chr(34)&"Screenshot"&chr(34)&")"


End Function


'''''''''''''''''''Reporting PASS''''''''''''''''''
Function reportPass(testName,testRowNum)
   ' first find the row from which test starts
 	Dim rNum,testStartRowNum
		rNum=1
		'Be careful
		while readData(environment("ResultxlsPath"),environment("testDataSheet"),rNum,1) <>  testName 
		rNum=rNum+1
		wend
	testStartRowNum = rNum
	resultColNum=1
	While readData(environment("ResultxlsPath"),environment("testDataSheet"),testStartRowNum+1,resultColNum) <> "Result"
		resultColNum =resultColNum+1
	Wend
	writeData environment("ResultxlsPath"),environment("testDataSheet"),testStartRowNum+1+testRowNum,resultColNum,"Pass"
End Function


'''''''''''''''''''Reporting Skip''''''''''''''''''
Function reportSkip(testName,testRowNum)
   ' first find the row from which test starts
		Dim rNum
		rNum=1
		'Be careful
		while readData(environment("ResultxlsPath"),environment("testDataSheet"),rNum,1) <>  testName 
		rNum=rNum+1
		wend
	
		testStartRowNum = rNum
	print testStartRowNum
	resultColNum=1
	While readData(environment("ResultxlsPath"),environment("testDataSheet"),testStartRowNum+1,resultColNum) <> "Result"
		resultColNum =resultColNum+1
	Wend
	writeData environment("ResultxlsPath"),environment("testDataSheet"),testStartRowNum+1+testRowNum,resultColNum,"Skipped"
End Function






