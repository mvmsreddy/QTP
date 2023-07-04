'Set box = Browser("creationtime:=0").Page("title:=.*").WebElement("html id:=frontbox")

Set d  = description.Create
d("micclass").value="WebElement"
d("html id").value="frontbox"

Set box = Browser("creationtime:=0").Page("title:=.*").ChildObjects(d)
print "Total such boxes - > "& box.count

Set desc= description.Create
desc("micclass").value="Link"

'Set box_links = box.ChildObjects(desc)
Set box_links = box(0).ChildObjects(desc)
print "Total links-> " & box_links.count

For i=0 to box_links.count-1
print box_links(i).getROProperty("innertext")
Next


