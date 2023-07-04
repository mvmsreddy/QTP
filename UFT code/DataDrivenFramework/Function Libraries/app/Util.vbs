
''''''''''''''''''''''''''''Logging in the app''''''''''''''''''''''''''''''
Function login(username,password)
   On error resume next
   LogThis "Logging in with "&username&"/"&password 
   ' check if user is already logged in or not
    If  isElementPresent("Logout") then
		'logout
		navigate environment("TestSiteURL")
	end if
'login the user
	findObject("LoginLink").click
	sync
	findObject("Username").set username
	findObject("Password").set password
	findObject("LoginButton").click
	sync
   'check if the user has logged in
   If  isElementPresent("Logout") then
	   login=true
	else
	  login=false
  end if
If err.number <> 0 Then
	err.clear
End If
End Function


'''''''''''''''''''''''''''''''''''''''''Default login'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function defaultLogin
   login environment("username"),environment("password")
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Double clicking on object ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function dbClickOnObject(obj)
Dim dRep,x,y
  Set dRep=createObject("Mercury.DeviceReplay")

  x= obj.getROProperty("abs_x")
  y= obj.getROProperty("abs_y")
dRep.MouseDblClick x+1,y+1,RIGHT_MOUSE_BUTTON

Set obj =nothing
Set dRep=nothing

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Select date'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function selectDate(dateTobeSelected)
Dim monthToBeSelected,yearToBeSelected,dayToBeSelected,desc,dates
monthToBeSelected  = monthName(month(dateTobeSelected),false)
yearToBeSelected = year(dateTobeSelected)
dayToBeSelected = cstr(day(dateTobeSelected))

print  monthToBeSelected
print yearToBeSelected
print dayToBeSelected

getPage().WebList("html id:=calMonthPicker","index:=0").Select monthToBeSelected
getPage().WebList("name:=calYearPicker","index:=0").Select cstr(yearToBeSelected)


Set desc= description.Create
desc("micclass").value="WebElement"
desc("class").value="weekday|weekend"

Set dates= getPage().ChildObjects(desc)
'print "total dates " & dates.count

For i=0 to dates.count-3
	If dates(i).getROProperty("innertext") =dayToBeSelected then
		dates(i).click
		Exit for
	end if
Next

Set desc=nothing
Set dates= nothing
End Function
			


