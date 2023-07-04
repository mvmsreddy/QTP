' HIDDEN OBJECTS
'Set objLink=Browser("creationtime:=0").Page("title:=.*").Link("name:=Business & Industrial")

Set objLink=Browser("creationtime:=0").Page("title:=.*").Link("name:=Customer Support","index:=0")

w = objLink.GetROProperty("width") 
h = objLink.GetROProperty("height")
x= objLink.GetROProperty("x")
y= objLink.GetROProperty("y")

print x
print y
print w
print h

If x=0 and y=0 and w=0 and h=0 Then
	print "Object is hidden"
else
    print "Object is visible"
End If


' DISABLED OBJECTS
Set mypage = Browser("creationtime:=0").Page("title:=.*")
print mypage.WebEdit("name:=realname","index:=0").GetROProperty("disabled")
print mypage.WebEdit("name:=realname","index:=0").GetROProperty("readonly")

print "-----------"
print mypage.WebEdit("name:=realname","index:=1").GetROProperty("disabled")
print mypage.WebEdit("name:=realname","index:=1").GetROProperty("readonly")


' Browser Functions
Browser("creationtime:=0").Back ' back button
wait(5)
Browser("creationtime:=0").Forward  ' forward

' Cache and cookies
Browser("creationtime:=0").ClearCache
Browser("creationtime:=0").DeleteCookies

WebUtil.DeleteCookies' delete cookies

' extract / read cookies

allCookies = Browser("creationtime:=0").Page("title:=.*").Object.cookie
print allCookies
temp=split(allCookies,";")
For i=0 to ubound(temp)
print temp(i)
Next

