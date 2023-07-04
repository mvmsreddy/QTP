
' export datatable
datatable.ExportSheet environment("Project_Path") & "results\Lead Suite - Results.xls","TestSteps"

LogThis "*********************ENDING Lead Suite****************"



'environment.LoadFromFile "E:\Tools\QTP\Recording Scripts\HybridFramework\env.xml"
'' import OR into datatable
'datatable.AddSheet("OR")
'datatable.ImportSheet "E:\Tools\QTP\Recording Scripts\HybridFramework\OR\OR.xls","OR","OR"
'' Dummy
''findObject("Username").set "ashish_thakur@sejsoft.com"
''findObject ("Password").set "pass@1234"
''findObject("LoginButton").click
'
'findObject("Username").set "ashish_thakur@sejsoft.com"
'findObjectByIndex("EditBox",1).set "hello"
'findObject("LoginButton").click
'
''closeAllBrowsers

'environment("Project_Path")="E:\Tools\QTP\Recording Scripts\HybridFramework\"
'environment("testSuiteName")="Lead Suite"
'Set data = readData ("CreateLead")
'print data.item(1).item("Phone")
'print data.item(2).item("Phone")
'print data.item(3).item("Phone")
'print data.item(4).item("Phone")
'
'readData "ConvertLead"
'readData "DeleteLead"


'Set d1 = CreateObject("Scripting.Dictionary")
'Set d2 = CreateObject("Scripting.Dictionary")
'Set d3 = CreateObject("Scripting.Dictionary")
'
'd1.Add "runmode","Y"
'd1.Add "Name" ,"x"
'
'd2.Add "runmode","N"
'd2.Add "Name" ,"y"
'
'd3.Add "runmode","Y"
'd3.Add "Name" ,"z"
'
'
'Set data = CreateObject("Scripting.Dictionary")
'data.Add 1,d1
'data.Add 2,d2
'data.Add 3,d3
'
'print data.Item(1).item("Name")
'
'print data.Item(3).item("runmode")

'
'environment.LoadFromFile "E:\Tools\QTP\Recording Scripts\HybridFramework\env.xml"
'		SystemUtil.Run environment("mozillaExePath")
'		sync
'
'	    SystemUtil.Run environment("ieExePath")
'		sync
'
'Browser("creationtime:=0").Navigate "gmail.com"
'
'print Browser("creationtime:=0").GetROProperty("application versione")


'Set desc = description.Create
'desc("micclass").value="WebElement"
'desc("html id").value="errorDiv_ep"
'Set obj = Browser("creationtime:=0").Page("title:=.*").ChildObjects(desc)
'print obj.count
'print obj(0).getROProperty("height")
'print obj(0).getROProperty("width")



'print replace(replace(replace(formatdatetime(date,1) &"_"& time,",","_"),":","_")," ","_")








