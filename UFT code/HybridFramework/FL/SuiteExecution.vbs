


Function executeTestSuite

testCasesRows = datatable.GetSheet("TestCases").GetRowCount
testStepsRows = datatable.GetSheet("TestSteps").GetRowCount
LogThis "Total test  cases Rows " &  testCasesRows
LogThis  "Total test  Steps Rows " &  testStepsRows


For tcid=1 to testCasesRows
	datatable.GetSheet("TestCases").SetCurrentRow(tcid)
	LogThis  "***************TestCase Under Execution -  " & datatable("TCID","TestCases") & " --- " & datatable("Runmode","TestCases") &"***********************"
	If  datatable("Runmode","TestCases") = "Y" Then
		' EXTRACT  DATA
		  currentTestCase = datatable("TCID","TestCases") 
		Set completeData = readData (currentTestCase)   ' dictionary
		For testRepeatNum=1 to completeData.count
			' make the result col
			resultColName="Result_"&testRepeatNum
			makeResultCol resultColName

		' execute the test
		If completeData.item(testRepeatNum).item("Runmode") ="Y"Then
		
           For tsid = 1 to testStepsRows
			   environment("errMsg")= ""
			 datatable.GetSheet("TestSteps").SetCurrentRow(tsid)
			   If  currentTestCase = datatable("TCID","TestSteps")  Then
					 keyword = datatable("Keyword","TestSteps")
					 objectName = datatable("Object","TestSteps")
					 dataKey = datatable("Data","TestSteps")
					 data  = completeData.item(testRepeatNum).item(dataKey)

						If  keyword = "openbrowser" Then
									result = openbrowser(data)
						elseif keyword = "click"  then
									result = click(objectName)
						elseif keyword = "navigate" then
									result = navigate(dataKey)
						elseif keyword = "input" then 
									result=input(objectName,data)
						elseif keyword = "selectList" then
									result=selectList(objectName,data)
						elseif keyword = "verifyTitle" then
									result=verifyTitle(dataKey)
						elseif keyword = "verifyElementPresence" then
									result=verifyElementPresence(objectName)
						elseif keyword = "verifyText" then
									result=verifyText(objectName,data)
						elseif keyword = "clickByIndex" then
									result = clickByIndex(objectName,datatable("Index","TestSteps"))
						elseif keyword = "verifyLeadCreation" then
									result = verifyLeadCreation(completeData.item(testRepeatNum).item("Flag"))
						elseif keyword = "defaultLogin" then
									result = defaultLogin(completeData.item(testRepeatNum))
						elseif keyword = "convertLead" then
									result = convertLead(completeData.item(testRepeatNum))
						elseif keyword = "login" then
									result = login(completeData.item(testRepeatNum))
						elseif keyword = "verifyLogin" then
									result = verifyLogin(completeData.item(testRepeatNum))	
						elseif keyword = "multiSelect" then
									result = multiSelect(objectName, data)					
						elseif keyword = "verifyListValues" then
									result = verifyListValues(objectName, data)	
						elseif keyword="verifyMenu" then
									result = 	verifyMenu(data)
						else

									msgbox "Keyword is not present in implementation " & keyword
									exittest  '' exit everything - critical
						End If

					 LogThis currentTestCase &" -- "& keyword &" --- "& data   &" ----- "& result

					' report error in xls - for every step
						 datatable(resultColName,"TestSteps") = result

					' take a screenshot
					If result <> "Pass" Then
						LogThis "ERROR for keyword  " & keyword &" -- " & result
						screenshotFileName = currentTestCase &"_" & keyword & "_"&testRepeatNum &".png"
						desktop.CaptureBitmap environment("Project_Path")&"results/screenshots/"&screenshotFileName,1
					End If
						'stop the current test
					 If result <> "Pass"  and datatable("Proceed_On_Failure","TestSteps") <> "Y" Then
						 Exit for ' come out of loop... stop executing keywords
					 End If
			   End If
	         
            Next
			else
			' Skip the data set
			print "Skipping the data set " & testRepeatNum
						 datatable(resultColName,"TestSteps") = "Skipped"
			End If
		   Next

	else
	' skip
	print "Skipping the test as Runmode is NO"
	' never reported here - skip   ----  YOU
	End If

Next






























End Function
