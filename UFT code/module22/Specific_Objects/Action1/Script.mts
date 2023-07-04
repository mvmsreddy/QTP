
' Description Obect

Set myPage = Browser("creationtime:=0").Page("title:=.*")
Set desc = description.Create  ' represents all the objects in all the applications
' Filter
desc("micclass").value="Link"

Set allPageObjects =    myPage.ChildObjects(desc)
print "Total objects found : "&  allPageObjects.count

' micclass
'print allPageObjects(0) illegal
print allPageObjects(0).getROProperty("micclass")  ''  br().page.Link().getroproperty


For i=0 to allPageObjects.count-1
  print allPageObjects(i).getroproperty("micclass") &" -- " & allPageObjects(i).getroproperty("innertext")       
Next

allPageObjects(0).select("Goa")













