Set table = Browser("creationtime:=0").Page("title:=.*").WebTable("class:=summary")

Set desc= description.Create
desc("micclass").value="Link"

Set allJobs =  table.ChildObjects(desc)

print allJobs.count



' Objects from specfic cells

Set table = Browser("creationtime:=0").Page("title:=.*").WebTable("index:=0")

' extract first button from cell 1,2
print "Total button objects in 1,2 -> " & table.ChildItemCount(1,2,"WebButton")
print "Total text   objects in 1,2 -> " & table.ChildItemCount(1,2,"WebEdit")
Set button = table.ChildItem ( 1,2, "WebButton", 1)
button.click






