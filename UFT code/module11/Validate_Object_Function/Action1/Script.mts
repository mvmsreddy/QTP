 'http://www.w3schools.com/html/default.asp
 bResult = bResult & "<table border=1><tr><th>" & _
                    "Status</th><th>Property</th><th>Expected</th><th>Actual</th></tr>"

Set allProp = Browser("HTML Tutorial").Page("HTML Tutorial").WebButton("Search").GetTOProperties
isPass=true
For i=0 to allProp.count -1
	expected = allProp(i).value
	actual = Browser("HTML Tutorial").Page("HTML Tutorial").WebButton("Search").GetROProperty(allProp(i).name)

	If trim(expected) <> trim(actual) Then
		bResult = bResult & "<tr><td>" & _
                            "Fail" & "</td><td>" & _
                            allProp(i).name & "</td><td>" & _ 
                            expected & "</td><td>" & _
                            actual & "</td></tr>"
		isPass=false
	else
	    bResult = bResult & "<tr><td>" & _
                            "Pass" & "</td><td>" & _
                            allProp(i).name & "</td><td>" & _ 
                            expected & "</td><td>" & _
                            actual & "</td></tr>"
	End If
Next



Set StepInfo = CreateObject("Scripting.Dictionary")
StepInfo("ViewType") = "Sell.Explorer.2"
StepInfo("DllIconIndex") = 209
StepInfo("DllIconSelIndex") = StepInfo("DllIconIndex")
StepInfo("DllPAth") = Environment.Value("ProductDir") & "\bin\ContextManager.dll"

If isPass Then
	StepInfo("Status") = micPass
else
    StepInfo("Status") = micFail
End If


StepInfo("NodeName") = "Result page name"

StepInfo("StepHtmlInfo") = "<html>" & bResult & "</body></html>"



    Reporter.LogEvent "User", StepInfo, Reporter.GetContext

   

'''''''''''''''''''''''''''''''''''''''''''''''''''FUNCTION''''''''''''''''''''''''''''''''''''''''''''

Set o =Browser("HTML Tutorial").Page("HTML Tutorial").WebButton("Search")
validateObjectProperties o,"search Button Validation"

Set o = Browser("HTML Tutorial").Page("HTML Tutorial").Link("HTML Formatting")
validateObjectProperties o,"Html Formatting Link"

Set x = Browser("The Times of India: Latest").Page("The Times of India: Latest").Frame("Frame").WebElement("1 + 5 =")
validateObjectProperties x,"TOI Link"



Function validateObjectProperties(object,stepName)

   bResult = bResult & "<table border=1><tr><th>" & _
                    "Status</th><th>Property</th><th>Expected</th><th>Actual</th></tr>"

Set allProp = object.GetTOProperties
isPass=true
For i=0 to allProp.count -1
	expected = allProp(i).value
	actual = object.GetROProperty(allProp(i).name)

	If trim(expected) <> trim(actual) Then
		bResult = bResult & "<tr><td>" & _
                            "Fail" & "</td><td>" & _
                            allProp(i).name & "</td><td>" & _ 
                            expected & "</td><td>" & _
                            actual & "</td></tr>"
		isPass=false
	else
	    bResult = bResult & "<tr><td>" & _
                            "Pass" & "</td><td>" & _
                            allProp(i).name & "</td><td>" & _ 
                            expected & "</td><td>" & _
                            actual & "</td></tr>"
	End If
Next



Set StepInfo = CreateObject("Scripting.Dictionary")
StepInfo("ViewType") = "Sell.Explorer.2"
StepInfo("DllIconIndex") = 209
StepInfo("DllIconSelIndex") = StepInfo("DllIconIndex")
StepInfo("DllPAth") = Environment.Value("ProductDir") & "\bin\ContextManager.dll"

If isPass Then
	StepInfo("Status") = micPass
else
    StepInfo("Status") = micFail
End If


StepInfo("NodeName") = stepName

StepInfo("StepHtmlInfo") = "<html>" & bResult & "</body></html>"

Reporter.LogEvent "User", StepInfo, Reporter.GetContext

 

End Function











