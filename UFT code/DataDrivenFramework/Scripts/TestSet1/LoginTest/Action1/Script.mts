environment.LoadFromFile environment("Project_Path")  & "env.xml"

'On error resume next
Dim testName,testData

testName ="LoginTest"
'Runmode of the test

'If NOT getRunmode(testName) Then
'	'report skip - YOU
'	LogThis "Skipping the test as runmode is NO"
'	destroyXLSApp
'	exittest
'End If

LogThis "Starting the test " & testName
' register the start time
registerTestCaseStartTime(testName)


Set testData = getData (environment("xlsPath") ,environment("testDataSheet"),testName)
LogThis "Total Rows are " & testData.count

For testIterator=1 to testData.count
environment("error")=""
' runmode for data set
If  testData.item(testIterator).item("Runmode") <> "Y" Then
	' report skipped test set

	reportSkip testName,testIterator
LogThis "Skipping  Login Test for iteration " & testIterator
else
environment("Browser")=testData.item(testIterator).item("Browser")
openBrowser 
sync
navigate environment("TestSiteURL")
result = login (testData.item(testIterator).item("Username"),testData.item(testIterator).item("Password"))
  If environment("error") <> "" Then
	' report error\
     reportError testName,testIterator,"Failed  Login Test for iteration. Err msg - " & environment("error")
     LogThis "Failed  Login Test for iteration. Err msg - " & environment("error")
  else	
	If NOT result and testData.item(testIterator).item("Flag") = "Y" Then
		'report error
		reportError testName,testIterator,"Failed  Login Test for iteration " & testIterator
		LogThis "Failed  Login Test for iteration " & testIterator
	elseif result and testData.item(testIterator).item("Flag") = "N" Then
		'report error
		reportError testName,testIterator,"Failed  Login Test for iteration " & testIterator
		LogThis "Failed  Login Test for iteration " & testIterator
	else
	     '  report pass
		 reportPass testName,testIterator
		LogThis "Passed  Login Test for iteration " & testIterator
	End If
 end if

End If

next
registerTestCaseEndTime(testName)
destroyXLSApp

'
'Set desc = description.Create
'desc("micclass").value="Browser"
'
'Set obj = desktop.ChildObjects(desc)
'msgbox obj.count
'
'msgbox obj(1).getROProperty("application version")
'
'
'
'print isBrowserOpen("Mozilla")
'print isBrowserOpen("internet explorer")
'

'msgbox chr(34)









