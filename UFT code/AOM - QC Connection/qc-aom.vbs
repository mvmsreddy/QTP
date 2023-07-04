 ' Connect to QC
Dim qtpqcApp
Set qtpqcApp= CreateObject("Quicktest.Application")
qtpqcApp.launch
qtpqcApp.visible=True

qtpqcApp.TDConnection.connect "http://localhost:8080/qcbin","HBAP","testproj","ashish","pass@123",False
msgbox qtpqcApp.TDConnection.IsConnected

' execute test
' open the test in QC
qtpqcApp.open "[QualityCenter] Subject\test\sampleTest" ' open project in QTP from QC
' add function lib
qtpqcApp.Test.Settings.Resources.Libraries.add "[QualityCenter\Resources] Resources\myTestResources\FL\fl"
' add OR
qtpqcApp.Test.Actions("Action1").ObjectRepositories.add "[QualityCenter\Resources] Resources\myTestResources\OR\object_rep"
qtpqcApp.Test.run
' disconnect
'qtpqcApp.TDConnection.Disconnect
'msgbox qtpqcApp.TDConnection.IsConnected
