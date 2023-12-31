
Function readData(testName)
   print "**************************"
Set objXls = createobject("Excel.application") 
objXls.Visible=false
Set myXls = objXls.Workbooks.Open (environment("Project_Path")&"Test Cases\"&environment("testSuiteName")&".xls")
Set sheet = myXls.Sheets("TestData")
 'find out the row from which the test starts
testStartRowNum=1
		while(sheet.cells(testStartRowNum,1).value <> testName)
			testStartRowNum=testStartRowNum+1			
		wend 
print "Test "&testName&" starts from row Num "& testStartRowNum
testDataStartRowNum = testStartRowNum+2
testColStartRowNum=testStartRowNum+1

' cols in the test case
totalCols=0
	    while(sheet.cells(testColStartRowNum,(totalCols+1)).value <> "")
			totalCols=totalCols+1
		wend

print "Test "&testName&" has total cols =  "& totalCols

'  rows in the test case
totalRows = 0
		while(sheet.cells(testDataStartRowNum+totalRows,1).value <>"")
		totalRows=totalRows+1
		wend

print "Test "&testName&" has total rows =  "& totalRows


' read the data

Set completeData = CreateObject("Scripting.Dictionary")
keyIndex=1
For rNum=testDataStartRowNum to testDataStartRowNum+totalRows-1
'data=""
Set rowData = CreateObject("Scripting.Dictionary")
	For cNum=1 to totalCols
		'data = data &" ---- " & sheet.cells(rNum,cNum)
		rowData.Add sheet.cells(testColStartRowNum,cNum).value,sheet.cells(rNum,cNum).value
	next
	completeData.Add keyIndex,rowData
	keyIndex=keyIndex+1
	'print data
Next


myXls.Close
objXls.Quit
Set objXls=nothing
Set myXls=nothing
Set sheet =nothing

Set readData=completeData

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
