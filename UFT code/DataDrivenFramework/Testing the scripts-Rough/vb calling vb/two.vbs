DIM objShell
set objShell = wscript.createObject("wscript.shell")
Dim fsObj
Set fsObj = CreateObject("Scripting.FileSystemObject")
Dim vbsFile
Set vbsFile = fsObj.OpenTextFile("E:\Tools\QTP\Recording Scripts\DataDrivenFramework\Testing the scripts-Rough\vb calling vb\three.vbs", 1, False)
Dim myFunctionsStr 
myFunctionsStr = vbsFile.ReadAll
msgbox myFunctionsStr
vbsFile.Close
Set vbsFile = Nothing
Set fsObj = Nothing
ExecuteGlobal myFunctionsStr 

sayHello

