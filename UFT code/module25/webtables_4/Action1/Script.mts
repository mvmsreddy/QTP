' http://www.espncricinfo.com/ci/engine/match/562896.html

Set table = Browser("creationtime:=0").Page("title:=.*").WebTable("html id:=inningsBat1")
print table.Exist
Set desc=description.Create
desc("micclass").value="WebElement"
desc("class").value="battingRuns"

Set runs = table.ChildObjects(desc)
total=0
For i=0 to runs.count-2
	total = total +  runs(i).getroproperty("innertext")
Next

print "total runs -> " & total







