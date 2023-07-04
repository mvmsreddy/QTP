Set desc = description.Create
desc("micclass").value="Window"
desc("text").value="Google Talk"

Set googleTalk= desktop.ChildObjects(desc)
print googleTalk.count

'  Winobject - contact list
desc("micclass").value="WinObject"
desc("text").value="Contact List"

Set contactList = googleTalk(0).childObjects(desc)
print contactList.count

'print contactList(0).GetVisibleText(3,144,250,180)  ' (l,t,r,b)
print contactList(0).GetVisibleText(3,180,250,220)  ' (l,t,r,b)
contactList(0).click 10,(180+220)/2


Set desc_chat=description.Create
desc_chat("micclass").value="Window"
desc_chat("regexpwndclass").value="Chat View"
Set chatWindow = desktop.ChildObjects(desc_chat)
print chatWindow.count



desc_chat("micclass").value="WinObject"
desc_chat("regexpwndclass").value="RichEdit20W"
Set msgB = chatWindow(0).childObjects(desc_chat)
print msgb.count
msgb(0).type "hello"


Set ws = createObject("WScript.shell")
ws.SendKeys "{Enter}"



' Your work
Function findnameAndClick(nameContact)

End Function



















