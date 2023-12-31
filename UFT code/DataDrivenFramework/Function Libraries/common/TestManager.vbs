'  find the object in the app and return it
Dim totalRows,ORFile,totalCols,objIndex
totalRows=0
objIndex=0
 Function findObject(ObjectName)
	Dim rNum,desc,obj,cNum,propertyVal,propertyType
  If  totalRows = 0 Then
	  ORFile=environment("Project_Path")&"OR\OR.xls"
		totalRows=getRowCount(ORFile,"OR")
		totalCols=getColumnCount(ORFile,"OR")
  End If

For rNum=2 to totalRows
	If  ObjectName = readData(ORFile,"OR",rNum,1)   Then
		' extract object from app
	Set desc=description.Create
		For cNum =2 to totalCols
			
			propertyVal=readData(ORFile,"OR",rNum,cNum)
				If  propertyVal <> "" Then
					propertyType=readData(ORFile,"OR",1,cNum)
					'print propertyType &" -- "& propertyVal
					desc(propertyType).value=propertyVal
				End If
		Next

			Set obj = getPage().childObjects(desc)
		
			If obj.count =0 Then
				environment("error")=environment("error") & vbnewline &"Object not found " & ObjectName
				Exit function
            ElseIf obj.count < objIndex Then
				environment("error")=environment("error") & vbnewline &"Object not found " & ObjectName
				Exit function
			else
				Set findObject = obj(objIndex) 
				Exit function
			End If
	End If

Next
'' Wrong Object name
msgbox "Object not Found in XLS " & ObjectName
exittest
 End Function


'' Finds the object by Index
Function findObjectByIndex(ObjectName,oIndex)
objIndex=oIndex
Set findObjectByIndex = findObject(ObjectName)
objIndex=0 'reset
End Function

''''''''''''''''''''''''''Return runmode status'''''''''''''''''''''''''''''''''
Function getRunMode(testCaseName)
' find total rows in the test cases sheet
dim tRows
tRows = getRowCount(environment("xlsPath"),environment("testCasesSheet"))

' iterate over every row and find the runmode
For rNum=2 to tRows
 If  testCaseName =  readData(environment("xlsPath"),environment("testCasesSheet"),rNum,1) Then

	If   readData(environment("xlsPath"),environment("testCasesSheet"),rNum,3) = "Y"  Then
		getRunMode=true
		Exit function
	else 
		getRunMode=false
		Exit function
	End If


 End If
Next

getRunMode=false
End Function

''''''''''''''''''''''''Finds the object on page'''''''''''''''''''''
Function isElementPresent(ObjectName)
   Dim rNum,desc,obj,cNum,propertyVal,propertyType
  If  totalRows = 0 Then
	  ORFile=environment("Project_Path")&"OR\OR.xls"
		totalRows=getRowCount(ORFile,"OR")
		totalCols=getColumnCount(ORFile,"OR")
  End If

For rNum=2 to totalRows
	If  ObjectName = readData(ORFile,"OR",rNum,1)   Then
		' extract object from app
	Set desc=description.Create
		For cNum =2 to totalCols
			
			propertyVal=readData(ORFile,"OR",rNum,cNum)
				If  propertyVal <> "" Then
					propertyType=readData(ORFile,"OR",1,cNum)
					'print propertyType &" -- "& propertyVal
					desc(propertyType).value=propertyVal
				End If
		Next

			Set obj = getPage().childObjects(desc)
		
			If obj.count =0 Then
				isElementPresent=false
				Exit function
			else
				isElementPresent=true
				Exit function
			End If
	End If

Next
'' Wrong Object name
msgbox "Object not Found in XLS " & ObjectName
exittest
End Function


Function registerTestCaseStartTime(testCaseName)
' find total rows in the test cases sheet
dim tRows
tRows = getRowCount(environment("xlsPath"),environment("testCasesSheet"))

' iterate over every row and find the runmode
For rNum=2 to tRows
 If  testCaseName =  readData(environment("xlsPath"),environment("testCasesSheet"),rNum,1) Then
	writeData environment("xlsPath"),environment("testCasesSheet"),rNum,4,now
 End If
Next
End Function


Function registerTestCaseEndTime(testCaseName)
   dim tRows
tRows = getRowCount(environment("xlsPath"),environment("testCasesSheet"))

' iterate over every row and find the runmode
For rNum=2 to tRows
 If  testCaseName =  readData(environment("xlsPath"),environment("testCasesSheet"),rNum,1) Then
	writeData environment("xlsPath"),environment("testCasesSheet"),rNum,5,now
 End If
Next
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Wait for element''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' true if element becomes visible in the waittime
' falso if not present till end of wait time
Function waitForElementPresence(obj,waittime)
  ' isElementPresent
End Function