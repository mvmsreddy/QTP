environment.LoadFromFile "E:\Tools\QTP\Recording Scripts\DataDrivenFramework\env.xml"

'openBrowser
'navigate("http://www.salesforce.com/in/?ir=1")

'setBrowserIndex(0)
'findObject("LoginLink",2).click
'findObject("Username").set "hello"
'findObject("Password").set "pass"
'findObject("LoginButton").click
'destroyFile

'findObjectByIndex("Objedit",0).set "hello"
'findObjectByIndex("Objedit",1).set "pass"
'
'destroyFile


environment("logFileName")= replace(replace(now,"/","_"),":","_")&".txt"
LogThis " starting test "
LogThis " executing test "
LogThis " ending test "






















