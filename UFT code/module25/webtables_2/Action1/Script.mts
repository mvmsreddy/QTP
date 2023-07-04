Set table = Browser("creationtime:=0").Page("title:=.*").WebTable("class:=oe-listview-content")
print table.GetROProperty("innertext")

print table.ChildItemCount(1,1,"WebButton")


print table.GetCellData(3,3)



table.ChildItem(4,9,"WebButton",0).click

