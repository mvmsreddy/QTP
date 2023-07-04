LogThis "*********************Starting Lead Suite****************"
environment.LoadFromFile environment("Project_Path") & "env.xml"
' Make sure all browsers are closed
closeAllBrowsers


'Set wsScript = CreateObject("Wscript.shell")
'msgbox wsScript.CurrentDirectory
'
datatable.AddSheet "TestCases"
datatable.AddSheet "TestSteps"
datatable.AddSheet("OR")
datatable.ImportSheet environment("Project_Path")&"OR\OR.xls","OR","OR"
datatable.ImportSheet environment("Project_Path") &"Test Cases\Lead Suite.xls","TestCases","TestCases"
datatable.ImportSheet environment("Project_Path") &"Test Cases\Lead Suite.xls","TestSteps","TestSteps"
environment("testSuiteName")="Lead Suite"

