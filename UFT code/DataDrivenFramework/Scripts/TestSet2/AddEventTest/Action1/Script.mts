environment.LoadFromFile environment("Project_Path")  & "env.xml"
copyResults  ''' later going in AOM
On error resume next
Dim testName,testData

testName ="AddEventTest"
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
	environement("error")=""
'check runmode of data set
	If  testData.item(testIterator).item("Runmode") <> "Y" Then
		LogThis "Skipping the iteration " & testIterator
	  reportSkip testName,testIterator
	else
	environment("Browser") = testData.item(testIterator).item("Browser")
	openbrowser
	' login if not logged in


	findObject("HomeTabLink").click
	sync
    findObject("EventTable").ChildItem(1,2,"Link",1).click
	sync
	findObject("NewEventButton").click
	sync
	findObject("EventSubject").set testData.item(testIterator).item("Subject")

	findObject("StartDateEvent").click
	selectDate testData.item(testIterator).item("StartDate")
	findObject("EndDateEvent").click
	selectDate testData.item(testIterator).item("EndDate")

	findobject("StartTimeEvent").set testData.item(testIterator).item("StartTime")
	findobject("EndTimeEvent").set testData.item(testIterator).item("EndTime")
	findObject("EventDescription").set testData.item(testIterator).item("Description")
	findObject("SaveProductButton").click
	sync
	findObject("DateImage").click

	selectDate testData.item(testIterator).item("StartDate")

	' do the verification  YOU

	reportPass testName,testIterator

	end if


next

destroyXlsApp
' rough








