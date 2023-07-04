text = browser("creationtime:=0").page("title:=Gmail: Email from Google").WebEdit("name:=Email").GetROProperty("value")
browser("creationtime:=0").page("title:=Gmail: Email from Google").WebEdit("name:=Email").Set text&"xxxxxxxx"
browser("creationtime:=0").page("title:=Gmail: Email from Google").WebEdit("name:=Passwd").Set "xxxx"

Set myapp=Dialog("regexpwndtitle:=Login")
myapp.Winedit("attached text:=Agent Name:").Set "mercury"
Dialog("regexpwndtitle:=Login").Winedit("attached text:=Password:").Set "mercury"
Dialog("regexpwndtitle:=Login").WinButton("text:=OK").Click


Set mypage = browser("creationtime:=0").page("title:=.*")
mypage.WebEdit("name:=Email","html id:=Email").Set text&"xxxxxxxx"
mypage.WebEdit("name:=Passwd").Set "xxxx"



Browser("creationtime:=1").Close
'  Tabs
'  No mixing of code
' Browsers


Browser("creationtime:=0").DeleteCookies
Set mypage = browser("creationtime:=0").page("title:=.*")
Browser("creationtime:=0").Navigate "http://demo.openerp.com/web/webclient/home"
mypage.webbutton("name:=Continue >").Click
mypage.webEdit("name:=dwf_email").Set "email@gmail.com"
mypage.webbutton("name:=Connect >").Click
mypage.Link("name:=  Sales ","index:=2").Click
wait(5)
mypage.WebButton("name:=Create","index:=0").Click
wait(5)
mypage.Image("file name:=down-arrow.png","index:=1").Click
mypage.WebElement("innertext:=Antony").Click
' find antony











