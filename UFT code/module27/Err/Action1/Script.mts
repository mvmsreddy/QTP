'Recovery Scenarios

On error resume next
Browser("Browser").Page("Gmail: Email from Google").WebEdit("Email").Set "ddd" @@ hightlight id_;_Browser("Browser").Page("Gmail: Email from Google").WebEdit("Email")_;_script infofile_;_ZIP::ssf3.xml_;_

If err.number <> 0 Then
	print err.number
	print err.description
End If
