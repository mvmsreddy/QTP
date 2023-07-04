
Set myPage = Browser("creationtime:=0").Page("title:=.*")
'myPage.Link("innertext:=2").Click

a=2
While myPage.Link("innertext:="&a).exist (10)
	myPage.Link("innertext:="&a).click
	myPage.Sync
	wait(5)
    printAllResults
	a=a+1
Wend


Function printAllResults

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


End Function

