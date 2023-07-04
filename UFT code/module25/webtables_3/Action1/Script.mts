' Dice.com
Set myPage=Browser("creationtime:=0").Page("title:=.*")
Dim a
a=2
While myPage.Link("innertext:="&a,"index:=0").exist
printJobs
myPage.Link("innertext:="&a,"index:=0").click
myPage.Sync
a=a+1
Wend

Function printJobs
Set table = Browser("creationtime:=0").Page("title:=.*").WebTable("class:=summary")

For i=1 to table.RowCount
print table.ChildItem(i,1,"Link",0).getROProperty("innerhtml")
Next
Set table = nothing
End Function



