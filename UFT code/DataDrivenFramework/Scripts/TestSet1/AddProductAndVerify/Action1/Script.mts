environment.LoadFromFile environment("Project_Path")  & "env.xml"
On error resume next
Dim testName,testData

testName ="AddProductAndVerify"

'Runmode of the test
'If NOT getRunmode(testName) Then
'	report skip - YOU
'	LogThis "Skipping the test as runmode is NO"
'	destroyXLSApp
'	exittest
'End If

Set testData = getData(environment("xlsPath"),environment("testDataSheet"),testName)
'msgbox testData.count


For testIterator=1 to testData.count
	environment("error")=""
	'check runmode of data set
	If  testData.item(testIterator).item("Runmode") <> "Y" Then
		LogThis "Skipping the iteration " & testIterator
	  reportSkip testName,testIterator
	else
	environment("Browser")=testData.item(testIterator).item("Browser")
	openBrowser
    ' check if logged in
	If NOT isElementPresent("Logout") Then
		navigate environment("TestSiteURL")
		defaultLogin
	End If

	findObject("ProductsTab").click
	sync
	findObject("NewProductButton").click
	findObject("ProductName").set testData.item(testIterator).item("ProductName")
	findObject("ProductCode").set testData.item(testIterator).item("ProductCode")
	findObject("ProductDescription").set testData.item(testIterator).item("Description")
	findObject("SaveProductButton").click
	sync

	If environment("error") <> "" Then ' some object is missing
		reportError testName,testIterator, environment("error")
	else
	' everything ok
	If  NOT verifyProperty("ProductNameDiplayed","innertext",testData.item(testIterator).item("ProductName")) then
		reportError testName,testIterator, " ProductNameDiplayed Property not matching"
	end if

	if  NOT verifyProperty("ProductCodeDisplayed","innertext",testData.item(testIterator).item("ProductCode")) then 
		reportError testName,testIterator, " ProductCodeDisplayed Property not matching"
	end if

	if  NOT verifyProperty("ProductDescriptionDisplayed","innertext",testData.item(testIterator).item("Description")) then 
		reportError testName,testIterator, " ProductDescriptionDisplayed Property not matching"
	end if

		findObject("ProductsTab").click
		sync
		findObjectByIndex("GoButton"  ,  1).click
		sync

dim prodName
prodName = testData.item(testIterator).item("ProductName")
Set desc = description.Create
desc("micclass").value="WebTable"
desc("class").value="x-grid3-row-table"

Set tables = Browser("creationtime:=0").Page("title:=.*").ChildObjects(desc)
print "Total tables are " & tables.count

		For i=0 to tables.count-1
			If  tables(i).getCellData(1,4) = prodName then
				tables(i).ChildItem (1, 3, "Link",1).click
			Exit for
			end if

		Next
        Set tables =nothing
		Set desc=nothing
		'' report pass
		reportPass "AddProductAndVerify",testIterator
End If
	
	End If



Next


Set testData = nothing


destroyXLSApp






'
'Dim prodName
'prodName=testData.item(testIterator).item("ProductName")
'Set desc = description.Create
'desc("micclass").value="WebElement"
'desc("class").value="x-grid3-cell-inner x-grid3-col-Name"
'
'Set obj = getPage().ChildObjects(desc)
'
'For i=0 to obj.count-1
' If obj(i).getROProperty("innertext") = prodName then
'	 ' activate the product
'	'Browser("creationtime:=0").Page("title:=.*").link("innertext:=Activate","index:="&i).click
'	findObjectByIndex("ActivateLink",i).click
'	Exit for
' end if
'Next











