environment.LoadFromFile environment("Project_Path")  & "env.xml"
copyResults  ''' later going in AOM
On error resume next
Dim testName,testData
Logthis "Add task test is executing"
testName ="AddTaskTest"
environment("xlsPath") = environment("Project_Path")&"Test Cases\TestSet2.xlsx"
environment("ResultxlsPath") = environment("Project_Path")&"Results\TestSet2.xlsx"
environment("logFilePath")= environment("Project_Path")&"logs/AddTaskTest.txt"
'msgbox getRunmode(testName)
'Runmode of the test
'If NOT getRunmode(testName) Then
'	'report skip - YOU
'	LogThis "Skipping the test as runmode is NO"
'	destroyXLSApp
'	exittest
'End If

' YOUR IMPLEMENTATION

 reportPass "AddTaskTest",1

destroyXLSApp
