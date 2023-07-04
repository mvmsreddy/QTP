dim strPath
msgbox "Google Test 1"
' load env var from QC
strPath=pathfinder.Locate("env")
print strPath
environment.LoadFromFile strPath
print environment("browser")

' import xls from QC
strPath=pathfinder.Locate("data")
datatable.AddSheet "my data"
datatable.ImportSheet strPath,"Test Data","my data"

print datatable("Search","my data") ' chocolate

'function lib
searchOnGoogle

'report fail
Reporter.ReportEvent micFail,"samplestep","xxxxxxxx"





Dim QCConnection 

Set QCConnection = QCUtil.QCConnection 

'Get the IBugFactory 

Set BugFactory = QCConnection.BugFactory 

'Add a new, empty defect 

Set Bug = BugFactory.AddItem (Nothing) 

'Enter values for required fields 

Bug.Status = "New" 

Bug.Summary = "New defect" 

Bug.DetectedBy = "admin" ' user that must exist in the database's user list 

'Post the bug to the database ( commit ) 

Bug.Post



