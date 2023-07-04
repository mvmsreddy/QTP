Set myPage = Browser("creationtime:=0").Page("title:=.*")
Set box =myPage.WebElement("html id:=news_moreTopStories")

Set desc = description.Create
desc("micclass").value="Link"

Set obj_link = box.ChildObjects(desc)
print "Total links - " & obj_link.count
'import

Dim total_links 
total_links =obj_link.count 

For i=0 to total_links-1
obj_link(i).click
myPage.Sync  


' http error
errMsg = validateURL()
  If errMsg <> "" then
	  ' capture screenshot
		Reporter.ReportEvent micFail,"bbc links","link not OK "& errMsg
  end if

' custom valiadation - tiles


Browser("creationtime:=0").Navigate "bbc.com"
'Browser("creationtime:=0").Back
'wait(5)

Set box =myPage.WebElement("html id:=news_moreTopStories")
Set obj_link = box.ChildObjects(desc)
Next




