' Google search result
Set myPage = Browser("creationtime:=0").Page("title:=.*")
Set desc= description.Create
desc("micclass").value="Link"
desc("class").value="l"
desc("height").value=18

Set objSearchLinks = myPage.ChildObjects(desc)
print objSearchLinks.count
For i=0 to objSearchLinks.count-1
print objSearchLinks(i).getROProperty("innertext")
Next


' demo.openerp.com
Set myPage = Browser("creationtime:=0").Page("title:=.*")
Set desc= description.Create
desc("micclass").value="WebButton"
desc("name").value="delete"
desc("disabled").value=0

Set objCrossButtons = myPage.ChildObjects(desc)

print objCrossButtons.count
objCrossButtons(3).click