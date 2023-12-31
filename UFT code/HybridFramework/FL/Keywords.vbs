Dim ieOpen,mozillaOpen
ieOpen = false
mozillaOpen = false

 Function openbrowser(bType)
		' if mozilla or IE  is opened or not
	If  mozillaOpen and bType = "Mozilla" Then
		openbrowser="Pass"
		setBrowserIndex(0)
		exit function
    elseif ieOpen and bType = "IE"  then
		openbrowser="Pass"
		setBrowserIndex(1)
		exit function
  end if


	If bType = "Mozilla" then
		SystemUtil.Run environment("mozillaExePath")
		mozillaOpen=true
		setBrowserIndex(0)
	elseif  bType = "IE"  then
	    SystemUtil.Run environment("ieExePath")
		setBrowserIndex(1)
		ieOpen=true
	end if
   sync  ' let the page load in browser

	openbrowser="Pass"
 End Function

 Function navigate(urlKey)
	Browser("creationtime:="&getBrowserIndex()).navigate environment(urlKey)
	navigate="Pass"
 End Function

 Function click(objectName)
	On error resume next
	findObject(objectName).click
	sync
    click="Pass"

	If err.number <>0 Then
	click=environment("errMsg")
	err.clear
	End If
 End Function

 Function selectList(objectName, selection)
	On error resume next
	findObject(objectName).select selection
	selectList="Pass"
    If err.number <>0 Then
	selectList=environment("errMsg")
	err.clear
	End If
 End Function



 Function input(objectName, data)
	On error resume next
	findObject(objectName).set data
	input="Pass"
	If err.number <>0 Then
	input=environment("errMsg")
	err.clear
	End If
 End Function

' verify title
  Function verifyTitle(expectedValueKey)
	 On error resume next
	actualTitle = getPage().getROProperty("title")
	expectedTitle = environment(expectedValueKey)
	If  actualTitle = expectedTitle Then
		verifyTitle="Pass"
	else
		verifyTitle= "Fail - titles did not match . Actual title - "& actualTitle
	End If
	If err.number <>0 Then
	input="Page not found" 
	err.clear
	End If
	
 End Function



 Function verifyElementPresence(objectName)
'	On error resume next
'	Set obj = findObject(objectName)
'	If obj is Nothing Then
'		verifyElementPresence = environment("errMsg")
'		environment("errMsg")=""
'	else
'		verifyElementPresence = "Pass"
'	End If

Set desc = description.Create

     For dRowNum=1 to datatable.GetSheet("OR").GetRowCount
		 datatable.GetSheet("OR").SetCurrentRow(dRowNum)
		 If  datatable("Name","OR") = objectName Then
			 Set desc = description.Create
				If datatable("micclass","OR")  <> "" then
					desc("micclass").value=datatable("micclass","OR")
				end if

				If datatable("html_id","OR")  <> "" then
					desc("html id").value=datatable("html_id","OR")
				end if

				If datatable("class","OR")  <> "" then
					desc("class").value=datatable("class","OR")
				end if

				If datatable("name","OR")  <> "" then
					desc("name").value=datatable("name","OR")
				end if

				If datatable("innertext","OR")  <> "" then
					desc("innertext").value=datatable("innertext","OR")
				end if
				
				If datatable("innerhtml","OR")  <> "" then
					desc("innerhtml").value=datatable("innerhtml","OR")
				end if

				If datatable("height","OR")  <> "" then
					desc("height").value=cint(datatable("height","OR"))
				end if

				Set obj = getPage().ChildObjects(desc)
				'msgbox obj.count
			
				' waiting for the object dynamically
				objectTimeout=0
				while(obj.count = 0  and objectTimeout < cint(environment("ObjectTimeOut"))  )
					wait(1)
					objectTimeout=objectTimeout+1
					Set obj = getPage().ChildObjects(desc)
				wend

				If obj.count = 0 Then
					' error
					verifyElementPresence="Fail -Object not found " & objectName
					'msgbox "object not found "+objectName
				elseif oIndex=0 then
					verifyElementPresence="Pass"
				else
					If cint(oIndex) >  obj.count Then
					verifyElementPresence="Fail -Object not found " & objectName
				    else
					verifyElementPresence="Pass"
					End If	
				End If
				Set desc=nothing
				Set obj =nothing
        End If
	 Next

 End Function



 Function verifyText(objectName,expectedText)
		On error resume next
		actualText = findObject(objectName).getROProperty("innertext")
		If  actualText = expectedText Then
			verifyText="Pass"
		else 
			verifyText = "Fail. Actual  text - " & actualText
		End If
		If err.number <>0 Then
			 verifyText= environment("errMsg") 
			err.clear
	    End If
 End Function

 Function clickByIndex(objectName,indexVal)
	On error resume next
	findObjectByIndex(objectName,indexVal).click
	clickByIndex="Pass"
	If err.number <>0 Then
	clickByIndex=environment("errMsg")
	err.clear
	End If

 End Function

 ''''''''''' Application dependent

 Function verifyLeadCreation(flag)  ' little flaw
	errMsgPresent = isHidden("LeadCreationErrMsg")
	print errMsgPresent
  If  NOT errMsgPresent and flag ="Y" Then
	  'error is present
    	environment("errMsg")=""  ''reset
	    verifyLeadCreation ="Fail - Lead could not be created with positive data"
  elseif   errMsgPresent and flag ="N" then
	    verifyLeadCreation ="Fail - Lead could  be created with negative data"
 else
	verifyLeadCreation ="Pass"
  End If

 End Function


 Function defaultLogin(testData)
		On error resume next
		If testData.item("Browser") ="Mozilla" Then
			setbrowserIndex(0)
		else
			setbrowserIndex(1)
		End If
		
	
		If  verifyElementPresence("LogoutLink") = "Pass"   Then
			defaultLogin="Pass"
			exit function
		End If

		'login
		openbrowser testData.item("Browser")
		navigate "testSiteBaseURL"
		click "LoginLink"
		input "Username",environment("defaultUsername")
		input "Password",environment("defaultPassword")
		click "LoginButton"
		defaultLogin="Pass"
		If err.number <>0 Then
		convertLead = environment("errMsg")
		err.clear
	    End If
 End Function

Function convertLead(testData)
   On error resume next
   click "LeadsLink"
   Browser("creationtime:="&getBrowserIndex).page("title:=.*").Link("innertext:="&testData.item("LeadName"),"index:=0").click
   sync
	click "LeadConvertButton"
	input "ConversionSubject",testData.item("Subject")
	selectList "ConversionStatus",testData.item("Status") 
	click "LeadConvertButton"
	convertLead="Pass"
	If err.number <>0 Then
		convertLead = environment("errMsg")
		err.clear
	End If
End Function


Function login(testData)
   On error resume next
		If testData.item("Browser") ="Mozilla" Then
			setbrowserIndex(0)
		else
			setbrowserIndex(1)
		End If
		
	
		If  verifyElementPresence("LogoutLink") = "Pass"   Then
			click "LogoutLink"

		End If

		'login
		openbrowser testData.item("Browser")
		navigate "testSiteBaseURL"
		click "LoginLink"
		input "Username",testData.item("Username")
		input "Password",testData.item("Password")
		click "LoginButton"
		login="Pass"
		If err.number <>0 Then
		login = environment("errMsg")
		err.clear
	    End If
 End Function

Function convertLead(testData)
   On error resume next
   click "LeadsLink"
   Browser("creationtime:="&getBrowserIndex).page("title:=.*").Link("innertext:="&testData.item("LeadName"),"index:=0").click
   sync
	click "LeadConvertButton"
	input "ConversionSubject",testData.item("Subject")
	selectList "ConversionStatus",testData.item("Status") 
	click "LeadConvertButton"
	convertLead="Pass"
	If err.number <>0 Then
		convertLead = environment("errMsg")
		err.clear
	End If
End Function

Function verifyLogin(testData)
   flag=testData("Flag")
   If  flag ="N" Then
			if  verifyText  ("LoginError",environment(testData.item("ExpectedError"))) ="Pass" then
							verifyLogin = "Pass"
			 else
							verifyLogin = "Login error message could not be verified"
			end if
   elseif flag="Y" Then
			If  verifyElementPresence("LoginError")  <> "Pass"  Then
							verifyLogin="Pass"
			else
								verifyLogin="Could not login"
			End If

    End If

'   If verifyElementPresence("LoginError") ="Pass" and flag ="N" then
'	  if  verifyText  ("LoginError",environment(testData.item("ExpectedError"))) ="Pass" then
'		verifyLogin = "Pass"
'	 else
'		verifyLogin = "Login error message could not be verified"
'	 end if
'   elseif verifyElementPresence("LoginError") ="Pass" and flag ="Y" then
'		verifyLogin = "Fail - not able to login with correct set of data"
'  elseif verifyElementPresence("LoginError") <> "Pass" and flag ="N" then
'		verifyLogin = "Fail - Able to login with wrong set of data"
'   elseif verifyElementPresence("LoginError") <> "Pass" and flag ="Y" then
'		verifyLogin = "Pass"
'   end if

End Function

' select multiple values from a list
Function multiSelect(object,data)
On error resume next
Dim arr,myList
Set myList = findObject(object)
arr = split(data,",")
myList.select arr(0)

For i=1 to ubound(arr)
myList.ExtendSelect arr(i)
Next
multiSelect="Pass"
Set myList = Nothing

if err.number <>0 Then
		multiSelect = environment("errMsg")
		err.clear
	End If
End Function


' Verify the values supplied in the list

Function verifyListValues(object,data)
On error resume next
Dim myList,cnt,temp,errMsg
errMsg=""
Set myList = findObject(object)
cnt = myList.getROProperty("items count")
temp=split(data,",")
	For i=0 to ubound(temp)
		'print "Finding "& temp(i) & "  in the list "
		For j=1 to cnt
			If  temp(i) = myList.getItem(j)   Then
			'	print "found"
				Exit for
			elseif j=cnt then
				errMsg = errMsg & "<" &temp(i) &" not found>"
			End If
		Next
	Next
If errMsg="" Then
	verifyListValues = "Pass"
else
    verifyListValues=errMsg
End If
if err.number <>0 Then
		verifyListValues = environment("errMsg")
		err.clear
	End If
End Function


'Check the menu for the links
Function verifyMenu(data)
On error resume next
Dim temp
verifyMenu=""
temp = split(data,",")

	For i =0 to ubound(temp)
		'print getPage().Link("innertext:="& temp(i)).exist
		If NOT getPage().Link("innertext:="& temp(i)).exist Then
			verifyMenu = verifyMenu &"<"& temp(i) &" not found >"
		End If
	Next
If  verifyMenu = "" Then
	verifyMenu="Pass"
End If

if err.number <>0 Then
		verifyListValues = environment("errMsg")
		err.clear
	End If
End Function






