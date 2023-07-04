' validate my username

' with OR
expectedValue="ashish thakur 1"
actualvalue=Browser("Rediff.com - India, Business,").Page("Rediffmail NG").Link("displayedName").GetROProperty("name")

If  expectedValue <> actualvalue Then
	Reporter.ReportEvent micFail,"Rediff Login",expectedValue &" ," & actualvalue
else
	Reporter.ReportEvent micPass,"Rediff Login",expectedValue &" ," & actualvalue
End If

If  Reporter.RunStatus <> micPass Then
 exitaction
End If

print "hello"
'' without OR
' YOUR WORK





