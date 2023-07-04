' Right click

' gaana.com
Setting.WebPackage("ReplayType")=2  ' analog

Set myPage = Browser("creationtime:=0").Page("title:=.*")

myPage.WebElement("class:=fl ellipsis","index:=0").FireEvent "onclick",1,1,micRightBtn

myPage.WebElement("class:=playnow").Click


' Keyboard input values
Set myPage = Browser("creationtime:=0").Page("title:=.*")
'myPage.WebElement("html id:=description1_ifr").Set "hello world"
'myPage.WebElement("html id:=description1_ifr").Click
myPage.WebElement("html id:=description1_ifr").FireEvent "onclick",1,1,micLeftBtn

Set o = CreateObject("wScript.shell")
o.SendKeys "hello world"


' Home Page
Browser("creationtime:=0").Home


' Tabs and poups - creationtime

Browser("creationtime:=0").Navigate "gmail.com"
Browser("creationtime:=1").Navigate "yahoo.com"

' moving mouse
wait(5)
Set o = CreateObject("Mercury.DeviceReplay")

'x=100
'y=200

x=Browser("creationtime:=0").Page("title:=.*").Link("name:=All Categories").GetROProperty("abs_x")
y=Browser("creationtime:=0").Page("title:=.*").Link("name:=All Categories").GetROProperty("abs_y")
o.MouseMove x,y





Services.StartTransaction "script_starting"

' Time taken by script to execute
st =  now
wait(5)
en= now

print datediff("s",st,en)
Reporter.ReportNote "Total time taken is " & datediff("s",st,en)

Services.EndTransaction "script_starting"





























