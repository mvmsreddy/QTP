Set myPage = Browser("creationtime:=0").Page("title:=.*")

Set desc = description.Create
desc("micclass").value ="WebTable"
desc("class").value ="dataTable"

Set table = myPage.ChildObjects(desc)
print table.count
'print table(0).getROProperty("innertext")


Set d = description.Create
d("micclass").value="Link"

Set objLinks = table(0).childObjects(d)
print objLinks.count

d("micclass").value="WebElement"
d("class").value="up"

Set objFigures = table(0).childObjects(d)
print objFigures.count

For i=0 to objLinks.count-1
	print objLinks(i).getROProperty("innertext") &" ---- " & objFigures(i).getROProperty("innertext")

Next
print table(0).getcelldata(1,1)







