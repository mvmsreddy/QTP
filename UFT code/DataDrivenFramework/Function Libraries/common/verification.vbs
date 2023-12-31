Option explicit
' true - object is enabled
 Function verifyEnabled(objectName)
	Dim obj,disabledProp
	Set obj = findObject(objectName)
	disabledProp= obj.getROProperty("disabled")

	If  disabledProp=0 Then ' enabled
		verifyEnabled=true
		 Reporter.ReportEvent micPass,"VerifyEnabled-"& objectName,objectName& " Object is enabled"
	else
		verifyEnabled=false
		' screenshot shot
		 Reporter.ReportEvent micFail,"VerifyEnabled-"& objectName,objectName& " Object is not enabled "
	End If

	Set obj=nothing
 End Function


Function verifyProperty(objectName,propName,expectedValue)
	Dim obj,actualValue
	Set obj = findObject(objectName)
	actualValue=obj.getROProperty(propName)

	If  actualValue=expectedValue Then ' Property matches
		verifyProperty=true
		 'Reporter.ReportEvent micPass,"VerifyProperty-"& objectName,"Property Matches"
	else
		verifyProperty=false
		' screenshot shot
		' Reporter.ReportEvent micFail,"VerifyProperty-"& objectName,"Expected <"&expectedValue&"> Actual<"&actualValue&">"
	End If
	Set obj=nothing
End Function

' finds if object is hidden or not
Function isHiddenObject(objectName)
  Dim obj,x,y,w,h
  Set obj = findObject(objectName)
  w=obj.getROProperty("width")
  h=obj.getROProperty("height")
  x=obj.getROProperty("x")
  y=obj.getROProperty("y")

  If x=0 and y=0 and w=0 and h=0 Then
	  isHiddenObject=true
  else
      isHiddenObject=false
  End If

	Set obj=nothing

End Function

' Hover mouse over any object
Function hovermouse(objectName)
  Dim obj,dRep,x,y
  Set obj = findObject(objectName)
  Set dRep=createObject("Mercury.DeviceReplay")

  x= obj.getROProperty("abs_x")
  y= obj.getROProperty("abs_y")
   dRep.MouseMove x,y

Set obj =nothing
Set dRep=nothing
End Function


' Hover mouse over any object
Function mouseDoubleClick(objectName)
  Dim obj,dRep,x,y
  Set obj = findObject(objectName)
  Set dRep=createObject("Mercury.DeviceReplay")

  x= obj.getROProperty("abs_x")
  y= obj.getROProperty("abs_y")
dRep.MouseDblClick x+1,y+1,RIGHT_MOUSE_BUTTON

Set obj =nothing
Set dRep=nothing
End Function
