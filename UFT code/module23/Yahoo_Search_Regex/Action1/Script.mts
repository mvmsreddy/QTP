Set desc = description.Create
desc("micclass").value="WebElement"
desc("html id").value="yui_3_4_0_1_1338923[0-9]{6}_[0-9]{4}"

Set obj = Browser("creationtime:=0").Page("title:=.*").ChildObjects(desc)
print obj.count

print obj(2).getRoproperty("innertext")

Set desc1 = description.Create
desc1("micclass").value="Link"

Set linkObj = obj(2).childObjects(desc1)
print linkObj.count
linkObj(2).click












