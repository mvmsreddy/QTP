Dim  browserIndex , oIndex
browserIndex=0  ' default
oIndex=0' default
'ieIndex=1
'mozillaIndex = 0


 Function findObject(objectName)

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
					raiseObjectNotFoundError(objectName)
					Set findObject = nothing
					'msgbox "object not found "+objectName
				elseif oIndex=0 then
					Set findObject = obj(0)
				else

					If cint(oIndex) >  obj.count Then
						print "here"
							raiseObjectNotFoundError(objectName+"<"&oIndex&">")
					       Set findObject = nothing 
				    else
					        Set findObject = obj(oIndex)   
					End If

					
				End If
				Set desc=nothing
				Set obj =nothing
        End If
	 Next

 End Function

' setting the browserIndex
 Function setBrowserIndex(bIndex)
	browserIndex=bIndex
 End Function

 ' getting the browserIndex
 function getBrowserIndex()
	getBrowserIndex = browserIndex
 end function


'  Page retrieving
Function getPage()
   Set getPage = Browser("creationtime:="&browserIndex).page("title:=.*")
End Function

Function sync
 getPage().sync
End Function


' close all browsers
Function closeAllBrowsers
   Set desc = description.Create
	desc("micclass").value="Browser"
	Set allBrowsers = desktop.ChildObjects(desc)
	For i=0 to allBrowsers.count-1
	    Browser("creationtime:="&i).close
	Next

Set desc=Nothing
Set allBrowsers =Nothing
SystemUtil.Run environment("mozillaExePath")
wait(2)
SystemUtil.Run environment("ieExePath")
ieOpen = true
mozillaOpen = true
End Function


'  multiple objects
Function findObjectByIndex(objectName,objectIndex)
oIndex = objectIndex
Set findObjectByIndex = findObject(objectName)
oIndex=0
End Function

Function raiseObjectNotFoundError(objectName)

If environment("errMsg") ="" Then
environment("errMsg") = "object is not found " & objectName
End If

'environment("errMsg") =environment("errMsg") &"---"&  "object is not found " & objectName
'msgbox environment("errMsg") 

End Function


 '' generate result col  -  testcasename_iteration
Function makeResultCol(colName)
On error resume next
datatable.GetSheet("TestSteps").GetParameter colName
If err.number <> 0 Then
	err.clear
    datatable.GetSheet("TestSteps").AddParameter colName," "
	Exit function
End If
End Function



Function isHidden(objectName)
	Set obj = findObject(objectName)
	If obj is nothing Then
			environment("errMsg")=""
			isHidden ="Object not found"
	End If

If obj.getROProperty("width") = 0 and obj.getROProperty("height") = 0 Then
		isHidden = true
else
		isHidden = false
End If
End Function

Function getCurrentFocusedBrowser
   If getBrowserIndex=0 Then
		getCurrentFocusedBrowser="Mozilla"
	else
		getCurrentFocusedBrowser="IE"
   End If
End Function





