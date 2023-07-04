' Mozilla - 3.5.x 
' IE = IE8
' QTP will  always record on browsers opened after opening QTP
'  Make sure that there are no add ins in your browser
' Windows 7 - turn off the UAC

SystemUtil.Run "H:\Program Files\Mozilla Firefox\firefox.exe","","H:\Program Files\Mozilla Firefox","open"
Browser("Mozilla Firefox Start").Page("Mozilla Firefox Start").Sync @@ hightlight id_;_Browser("Mozilla Firefox Start").Page("Mozilla Firefox Start")_;_script infofile_;_ZIP::ssf1.xml_;_
Browser("Mozilla Firefox Start").navigate "gmail.com"
Browser("Mozilla Firefox Start").Page("Gmail: Email from Google").WebEdit("Email").Set "xxxxx" @@ hightlight id_;_Browser("Mozilla Firefox Start").Page("Gmail: Email from Google").WebEdit("Email")_;_script infofile_;_ZIP::ssf2.xml_;_
Browser("Mozilla Firefox Start").Page("Gmail: Email from Google").WebEdit("Passwd").SetSecure "4ecea57983b7cc6235b14a97e09affe5536c0e47" @@ hightlight id_;_Browser("Mozilla Firefox Start").Page("Gmail: Email from Google").WebEdit("Passwd")_;_script infofile_;_ZIP::ssf3.xml_;_
Browser("Mozilla Firefox Start").Page("Gmail: Email from Google").WebButton("Sign in").Click @@ hightlight id_;_Browser("Mozilla Firefox Start").Page("Gmail: Email from Google").WebButton("Sign in")_;_script infofile_;_ZIP::ssf4.xml_;_


' WINDOWS BASED APP
' Record and run settings
' Only one instance of the application should be running


SystemUtil.Run "calc","","H:\Documents and Settings\Administrator",""
Window("Calculator").WinButton("9").Click @@ hightlight id_;_723522_;_script infofile_;_ZIP::ssf5.xml_;_
Window("Calculator").WinButton("6").Click @@ hightlight id_;_723390_;_script infofile_;_ZIP::ssf6.xml_;_
Window("Calculator").WinButton("3").Click @@ hightlight id_;_723422_;_script infofile_;_ZIP::ssf7.xml_;_
Window("Calculator").WinButton("2").Click @@ hightlight id_;_788924_;_script infofile_;_ZIP::ssf8.xml_;_
Window("Calculator").Click 91,101 @@ hightlight id_;_984392_;_script infofile_;_ZIP::ssf9.xml_;_
Window("Calculator").WinButton("5").Click @@ hightlight id_;_854444_;_script infofile_;_ZIP::ssf10.xml_;_

' Windows 7
'Systemutil.Run "H:\Program Files\Mozilla Firefox\firefox.exe","cnn.com"
Browser("Mozilla Firefox Start").Page("Mozilla Firefox Start").Sync @@ hightlight id_;_Browser("Mozilla Firefox Start").Page("Mozilla Firefox Start")_;_script infofile_;_ZIP::ssf11.xml_;_
Browser("Mozilla Firefox Start").navigate "qtpselenium.com"
Browser("Mozilla Firefox Start").Page("Online Selenium Training").Image("view_details").Click @@ hightlight id_;_Browser("Mozilla Firefox Start").Page("Online Selenium Training").Image("view details")_;_script infofile_;_ZIP::ssf12.xml_;_

Browser("Mozilla Firefox Start").Page("Yahoo! India").Link("Mail").Click @@ hightlight id_;_Browser("Mozilla Firefox Start").Page("Yahoo! India").Link("Mail")_;_script infofile_;_ZIP::ssf13.xml_;_




SystemUtil.Run "H:\Program Files\Mozilla Firefox\firefox.exe","","H:\Program Files\Mozilla Firefox","open"
Browser("Mozilla Firefox Start").Page("Mozilla Firefox Start").Sync @@ hightlight id_;_Browser("Mozilla Firefox Start").Page("Mozilla Firefox Start")_;_script infofile_;_ZIP::ssf14.xml_;_
Browser("Mozilla Firefox Start").navigate "http://way2sms.com/"
optionalstep.Browser("Mozilla Firefox Start").Page("Way2SMS - Send Free SMS").Link("click here to go to way2sms.co").Click @@ hightlight id_;_Browser("Mozilla Firefox Start").Page("Way2SMS - Send Free SMS").Link("click here to go to way2sms.co")_;_script infofile_;_ZIP::ssf15.xml_;_



'   Password Encoding 
Browser("Mozilla Firefox Start").Page("Gmail: Email from Google").WebEdit("Email").Setsecure "4eceaf9214513cc7df4457bfe8049b47167a3edf19b9dcd7fba4" @@ hightlight id_;_Browser("Mozilla Firefox Start").Page("Gmail: Email from Google").WebEdit("Email")_;_script infofile_;_ZIP::ssf2.xml_;_



