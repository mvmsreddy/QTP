Dim qtpApp,xlsApp,myFile,suiteSheet,rNum,testSuiteName,projectPath,logFilePath
Const testSuiteSheet = "TestSuites"


' remove harcoding of paths
' VB to find the current Directory location

Set ws = CreateObject("Wscript.Shell")
projectPath =  ws.CurrentDirectory & "\"
'msgbox projectPath
'projectPath="E:\Tools\QTP\Recording Scripts\HybridFramework\"


' generate log file path
logFilePath = projectPath &"logs\"&replace(replace(replace(formatdatetime(date,1) &"_"& time,",","_"),":","_")," ","_")&".txt"





' open QTP
openQTP

'check the runmodes of test Suites
executeScript
' execute the test suite scripts

' screenshots of errors

' Logging in a txt file




Function openQTP
Set qtpApp = CreateObject("QuickTest.Application")
qtpApp.Visible=True

	If NOT qtpApp.Launched Then
			qtpApp.launch
	End if
End function


Function executeScript
Set xlsApp = CreateObject("Excel.Application")
xlsApp.visible=false
Set myFile =  xlsApp.Workbooks.Open(projectPath&"Test Cases\TestSuite.xlsx")
Set suiteSheet = myFile.sheets(testSuiteSheet)
rNum=2

While suiteSheet.cells(rNum,1) <> "" 
		testSuiteName = suiteSheet.cells(rNum,1)
		If suiteSheet.cells(rNum,2) = "Y" Then
			loadAndRunTestSuite  ' calling for execution of the test Suite
		End if
		rNum = rNum + 1
Wend 


myFile.close
xlsApp.quit

Set xlsApp = nothing
Set myFile = Nothing
Set suiteSheet = Nothing
End Function


Function loadAndRunTestSuite

' open the driverscript corresponding to test suite
msgbox projectPath &"drivers\Driver - "&testSuiteName
qtpApp.open projectPath &"drivers\Driver - "&testSuiteName

' Set the path env var in the test
qtpApp.Test.Environment("Project_Path") = projectPath
qtpApp.Test.Environment("logFilePath") = logFilePath

qtpApp.Test.Run

End Function


