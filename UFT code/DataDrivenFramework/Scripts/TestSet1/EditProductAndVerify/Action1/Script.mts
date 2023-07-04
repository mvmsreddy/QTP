'http://go-gaga-over-testing.blogspot.in/2011/07/maximize-browser-using-qtp-and-ie8-ie7.html

environment.LoadFromFile environment("Project_Path")  & "env.xml"
'On error resume next
Dim testName,testData

testName ="EditProductAndVerify"

'Runmode of the test

'If NOT getRunmode(testName) Then
'	'report skip - YOU
'	LogThis "Skipping the test as runmode is NO"
'	destroyXLSApp
'	exittest
'End If

LogThis "Starting the test " & testName
Set testData = getData(environment("xlsPath"),environment("testDataSheet"),testName)
'msgbox testData.count


For testIterator=1 to testData.count
	' RUnmode for the dataset
	'check runmode of data set
If  testData.item(testIterator).item("Runmode") <> "Y" Then
	LogThis "Skipping the iteration " & testIterator
	  reportSkip testName,testIterator
else
	environment("error")=""
	Environment("Browser") = testData.item(testIterator).item("Browser")
	openBrowser
	' login and reach the product page -YOU
	LogThis  "Expecting iteration "& testIterator


	Set desc = description.Create
	desc("micclass").value="WebTable"
	desc("class").value="x-grid3-row-table"

	Set tables = getPage.childObjects(desc)
	LogThis  "Total tables "& tables.count
	prodName = testData.item(testIterator).item("ProductName")

'Browser("creationtime:=0").FullScreen
wait(2)
	For i=0 to tables.count-1
		If  tables(i).getCellData(1,4) =  prodName   Then
			LogThis "Found " & i
			' editing prod name			
			If testData.item(testIterator).item("NewProductName") <> "" Then	
			dbClickOnObject tables(i).ChildItem(1,4,"WebElement",0)
			findObject("ProductName").set testData.item(testIterator).item("NewProductName") 
            findObject("SaveProductButton").click
			sync
			End If
			wait(2)
			' editing prod code			
			If testData.item(testIterator).item("NewProductCode") <> "" Then			
			dbClickOnObject tables(i).ChildItem(1,5,"WebElement",0)
			findObject("ProductCode").set testData.item(testIterator).item("NewProductCode") 
			findObject("SaveProductButton").click
			sync
			End If
			wait(2)
			' editing prod code			
			If testData.item(testIterator).item("NewDescription") <> "" Then			
			dbClickOnObject tables(i).ChildItem(1,6,"WebElement",0)
			findObject("ProductDescription").set testData.item(testIterator).item("NewDescription") 
			findObject("SaveProductButton").click
			sync
			End If
			
		  reportPass testName,testIterator
			Exit for
		End If
		' product is not found
		If i = tables.count-1Then
			reportError testName,testIterator, " Product not found" 
		End If

	Next
	Set desc= nothing
	Set tables = nothing

end if
next


destroyXLSApp


































