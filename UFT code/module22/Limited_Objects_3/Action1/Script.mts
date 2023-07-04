Set desc = description.Create
desc("micclass").value="WebList"
desc("name").value="select"

Set objList = Browser("creationtime:=0").Page("title:=.*").ChildObjects(desc)
print objList.count
wait(5)
objList(0).CaptureBitmap "c:\list.png", true
print objList(0).GetItem (2)  ' 1
print objList(0).select (2)  ' 0


 'Browser("qtp Jobs at Dice.com").Page("qtp Jobs at Dice.com").WebList("select").CaptureBitmap




