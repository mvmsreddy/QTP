Set desc = description.Create
desc("micclass").value="WebElement"
desc("innerhtml").value="  Gmail is built on the.*"

Set obj = Browser("creationtime:=0").Page("title:=.*").ChildObjects(desc)
print obj.count



print Browser("creationtime:=0").Page("title:=.*").WebElement("innerhtml:=Gmail is built on the.*").exist