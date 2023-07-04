Dim qtpApp,logFileName
' init logfile
logfileName = replace(replace(now,"/","_"),":","_")&".txt"
Set qtpApp=createObject("Quicktest.Application")
qtpApp.Visible=true

If  not qtpApp.Launched  Then
	' launch qtp
qtpApp.Launch
End If
LogThis "QTP opened"

qtpApp.options.run.RunMode="Fast"


' ************************Modified datatable**************************
'qtpApp.test.datatable.addSheet "TestSheet"
'qtpApp.test.datatable.getSheet("TestSheet").AddParameter "Country","India"
'qtpApp.test.datatable.exportSheet "C:\temp\aom.xls","TestSheet"

'************************Open test and Run***********************
LogThis "executing first test"

qtpApp.Open "E:\Tools\QTP\Recording Scripts\QTP AOM\Yahoo"
qtpApp.Test.Actions("Action1").ObjectRepositories.add "E:\Tools\QTP\Recording Scripts\QTP AOM\yahoorep.tsr"
qtpApp.Test.Settings.Resources.Libraries.add "E:\Tools\QTP\Recording Scripts\QTP AOM\Lib.vbs"
qtpApp.Test.Run
LogThis "Test 1 - " & qtpApp.test.LastRunResults.Status
qtpApp.Test.close



LogThis "executing second test"
qtpApp.Open "E:\Tools\QTP\Recording Scripts\QTP AOM\Yahoo"
qtpApp.Test.Actions("Action1").ObjectRepositories.add "E:\Tools\QTP\Recording Scripts\QTP AOM\yahoorep.tsr"
qtpApp.Test.Settings.Resources.Libraries.add "E:\Tools\QTP\Recording Scripts\QTP AOM\Lib.vbs"
qtpApp.Test.Run
LogThis "Test 2 - " & qtpApp.test.LastRunResults.Status
qtpApp.Test.close

LogThis " Tests Executed "

'WScript.sleep(5000)
'qtpApp.Quit

'Set qtpApp=nothing


Function LogThis(msg)
Dim fso,myfile
Const ForAppening=8
Set fso=CreateObject("Scripting.FileSystemObject")
Set myfile=fso.OpenTextFile("E:\Tools\QTP\Recording Scripts\QTP AOM\logs"&logfileName,ForAppening,true)
myfile.writeLine now &"--" &msg

Set fso=Nothing
Set myfile=nothing
End function

