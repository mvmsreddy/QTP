print parameter("username")
print parameter("password")

Browser("Rediff.com - India, Business,").Navigate "in.rediff.com"
Browser("Rediff.com - India, Business,").Page("Rediff.com - India, Business,").Link("Sign in").Click
Browser("Rediff.com - India, Business,").Page("Rediff.com - India, Business,").WebEdit("id").Set parameter("username")
Browser("Rediff.com - India, Business,").Page("Rediff.com - India, Business,").WebEdit("num").Set parameter("password")
Browser("Rediff.com - India, Business,").Page("Rediff.com - India, Business,").WebButton("Go").Click

' return parameter - login success
 If Browser("Rediff.com - India, Business,").Page("Rediffmail NG").Link("Inbox").Exist(20) then
	 ' return login success
	 parameter("loginSuccess")=true
 else
      'return login failure
	 parameter("loginSuccess")=false
 end if
