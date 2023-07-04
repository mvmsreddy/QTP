' Yahoo Search

Set myPage=Browser("creationtime:=0").Page("title:=.*")
Set desc = description.Create
desc("micclass").value="WebEdit"
Set objWebEdit = myPage.ChildObjects(desc)
print "Total edit boxes "& objWebEdit.count

objWebEdit(0).set "top one"
objWebEdit(1).set "bottom one"

' destroy objects
Set myPage= nothing
Set desc= nothing
Set objWebEdit = nothing

'  demo.virtuemart.com

Set myPage=Browser("creationtime:=0").Page("title:=.*")
'myPage.Link("class:=product-name","index:=1").Click
Set desc = description.Create
desc("micclass").value="Link"
desc("class").value="product-name"
Set objProds = myPage.ChildObjects(desc)
print "Total Products -> "& objProds.count
objProds(1).click '' clicks on second link



''  All product detail buttons
Set myPage=Browser("creationtime:=0").Page("title:=.*")
'myPage.Link("class:=product-name","index:=1").Click
Set desc = description.Create
desc("micclass").value="Link"
desc("name").value="Product Details"
Set objProds = myPage.ChildObjects(desc)
print "Total Buttons -> "& objProds.count

'myPage.Link("name:=Product Details","index:=3").Click

Set descProdNames = description.Create
descProdNames("micclass").value="WebElement"
descProdNames("class").value="width18 floatleft"
Set objProdNames=myPage.ChildObjects(descProdNames)
print "Total Products -> "& objProdNames.count

Dim prdName

For i=0 to objProdNames.count-1
prdName =  objProdNames(i).getroproperty("innertext")
	If  trim(prdName) = "iPhone 4" Then
		objProds(i).click
	End If
Next













