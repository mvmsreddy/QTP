datatable.AddSheet "quikr"

set myPage = Browser("creationtime:=0").Page("title:=.*")

Set descCategories = description.Create
descCategories("micclass").value="WebElement"
descCategories("html id").value = "cats"

Set categories= myPage.ChildObjects(descCategories)
print categories.count

Set desc_link = description.Create
desc_link("micclass").value= "Link"

Dim col_name
' category
For i=0 to categories.count-1
	col_name = "Category_"& i
datatable.GetSheet("quikr").AddParameter col_name,""
	print "***********************************************"
Set objLink = categories(i).childObjects(desc_link)
' links
print "Total Links -> " & objLink.count
	For j=0 to objLink.count-1
		datatable.GetSheet("quikr").SetCurrentRow(j+1)
		linktext= objLink(j).getROProperty("innertext")	
		print linktext
		datatable(col_name,"quikr")=linktext
	Next
 
Next

datatable.ExportSheet "C:\quikr.xls","quikr"




