Set box = Browser("creationtime:=0").Page("title:=.*").WebElement("class:=signin-box")

Set desc = description.Create
desc("micclass").value="WebEdit"

Set objEdits = box.ChildObjects(desc)
print " Total edit boxes -> "& objEdits.count
objEdits(0).set "username"
objEdits(1).set "password"



desc("micclass").value="WebButton"
Set objButton = box.ChildObjects(desc)
print "Total buttons -> " & objButton.count
print objButton(0).getroproperty("value")
'objButton(0).click


desc("micclass").value="WebCheckBox"
Set objCheckBox = box.ChildObjects(desc)
If  objCheckBox(0).getROProperty("checked") = 0 Then
objCheckBox(0).click
End If



