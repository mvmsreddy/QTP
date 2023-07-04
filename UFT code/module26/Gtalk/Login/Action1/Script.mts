
print Window("text:=Google Talk").exist

Set desc = description.Create
desc("micclass").value="Window"
desc("text").value="Google Talk"

Set googleTalk= desktop.ChildObjects(desc)
print googleTalk.count
'
' enter username and password
Set desc_edit=description.Create
desc_edit("micclass").value="WinEdit"
Set inputFields = googleTalk(0).childObjects(desc_edit)
print inputFields.count
inputFields(0).set "qtpselenium@gmail.com"
inputFields(1).set "whizdomtrainings"

' button
desc_edit("micclass").value="WinButton"
Set button = googleTalk(0).childObjects(desc_edit)
button(0).click





