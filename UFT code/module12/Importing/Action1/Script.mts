datatable.AddSheet("testData")
datatable.Import "C:\qtp.xls"



datatable.GetSheet("Action1").SetCurrentRow(4)
print datatable("Country","Action1")



' import a sheet testData
datatable.AddSheet("currentTest")
datatable.ImportSheet "C:\qtp.xls","testData","currentTest"
datatable.GetSheet("currentTest").SetCurrentRow(5)
print datatable("Test","currentTest")