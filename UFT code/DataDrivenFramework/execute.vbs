
 Dim  currentDir,ws,qtpApp,testSetXlsPath,testSets,xlApp,myXls,sheet,rowNumber,currentTestSetName,currentTestCase,logFileName,logFilePath
 Const testCasesFolder = "Test Cases"
 Const testSetXLSName = "TestSets.xlsx"
 Const testSetSheet="TestSets"
 Const testCasesSheet="Test Cases"
 Const scriptsFolderName="Scripts"
 Const resultFolderName="Results"
 Const logFolder="logs"
 ' generate LogFile
generateLogFile
 ' Find the current directory
Set ws  = createobject("Wscript.shell")
currentDir =  ws.CurrentDirectory

LogThis "Starting AOM Script"
 
 
' Open up QTP
openQTP
LogThis "QTP opened"
copyResults



'Begin Execution
testSetXlsPath = currentDir & "\" & testCasesFolder & "\" & testSetXLSName
'msgbox testSetXlsPath


Set xlApp= CreateObject("Excel.Application")
xlApp.Visible=false
xlApp.DisplayAlerts=False
Set myXls = xlApp.Workbooks.Open(testSetXlsPath)
Set sheet=myXls.sheets(testSetSheet)
rowNumber=2

While sheet.cells(rowNumber,1) <> ""
'msgbox sheet.cells(rowNumber,1) & "  ---  " &  sheet.cells(rowNumber,2)
If sheet.cells(rowNumber,2) = "Y" Then
currentTestSetName = sheet.cells(rowNumber,1)
executeTestSet currentTestSetName
End if
rowNumber=rowNumber+1
Wend

 myXLS.save
 myxls.close
 xlApp.Application.Quit
Set xlApp=Nothing
Set myXls=Nothing
Set sheet=nothing


'closeQTP
LogThis "AOM Script finishes"
msgbox "Script over"

'''''''''''''''''''''Opens QTP'''''''''''''''''''''''
Function openQTP
Set qtpApp = CreateObject("QuickTest.Application")
qtpApp.visible=false

If Not qtpApp.launched Then
qtpApp.launch
End if
'qtpApp.Options.Run.Runmode="Fast"
End Function

''''''''''''''''''''''''Executing the test Cases in a test Set'''''''''''''''''''''''''''''
Function executeTestSet(testSetName)
'msgbox "Executing " & testSetName
Set testSetXls = xlApp.Workbooks.Open(currentDir & "\" & testCasesFolder & "\" &testSetName & ".xlsx")
Set casesSheet = testSetXls.sheets(testCasesSheet)

row=2
While casesSheet.cells(row,1) <> ""
currentTestCase = casesSheet.cells(row,1)
executeTestCase currentTestCase
row=row+1
Wend


testSetXls.save
testSetXls.close
Set casesSheet =Nothing
Set testSetXls = nothing

End Function



''''''''''''''''''''''Executing the test Cases''''''''''''''''''''''''''''
Function executeTestCase(testCaseName)
'msgbox currentTestSetName &" -- " & testCaseName
Dim runmode
' check runmode and report skip if N - YOU
runmode="Y"


If runmode = "Y" then
LogThis "EXECUTING TEST CASE " & testCaseName & "IN TESTSET "& currentTestSetName
qtpApp.open currentDir & "/" & scriptsFolderName & "/" & currentTestSetName & "/" & testCaseName
' setting env variables
qtpApp.Test.Environment("Project_Path") = currentDir &"/"
qtpApp.Test.Environment("xlsPath") = currentDir & "/" & testCasesFolder & "/" & currentTestSetName& ".xlsx"
qtpApp.Test.Environment("ResultxlsPath") = currentDir & "/" & resultFolderName & "/" & currentTestSetName& ".xlsx"
'qtpApp.Test.Environment("logFilePath")= currentDir & "/" & logFolder & "/" & testCaseName & ".txt"
qtpApp.Test.Environment("logFilePath")= currentDir & "/" & logFolder & "/"  & logFileName
qtpApp.Test.run
Else
' REPORT the skipping of all test sets of the test case - YOU
End if


End Function
' RUNMODE for a test case
Function getTestRunmode(testName)
' YOUR
End function

''''''''''''''''Make result folder'''''''''''''''''''''''
''''''''''''''''Copying the testset files in the result folder'''''''''''''  - later going inside AOM
Function copyResults
Dim fso,folderHandle,fileName,src,dest
Set fso = CreateObject("Scripting.FileSystemObject")
Set folderHandle = fso.GetFolder(currentDir & "/" & testCasesFolder)
Set allFiles = folderHandle.files
For Each f1 in allFiles
fileName = f1.name 
src = currentDir & "/" & testCasesFolder & "/" & fileName
dest = currentDir & "/" & resultFolderName & "/" & fileName
fso.CopyFile  src,dest,True
Next
Set fso = nothing
Set folderHandle = nothing
Set allFiles = nothing
End Function



'''''''''''''''''''''Generating log file name''''''''''''''''''''''''''''
Function generateLogFile
logFileName = replace(replace(FormatDateTime(Date, 1) &"_" & time,",","_"),":","_") &".txt"
End Function

'''''''''''''''''''''Logging''''''''''''''''''''''''''
 
Function LogThis(msg)
   If logFilePath="" Then
	   logFilePath= currentDir & "/" & logFolder & "/"  & logFileName
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


'close QTP
Function closeQTP
qtpApp.Test.Close
qtpApp.Quit
Set qtpApp = nothing
End function
