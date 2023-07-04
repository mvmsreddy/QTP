' Local Sheet

print "Total row - " & datatable.GetSheet("Action1").GetRowCount 

For i=1 to datatable.GetSheet("Action1").GetRowCount 
datatable.GetSheet("Action1").SetCurrentRow(i)
print datatable("Country","Action1")
datatable("Test","Action1")=i
Next

'Example

datatable.GetSheet("Action1").AddParameter "Col4","123"
' Add own sheet in datatable
datatable.AddSheet("testData")
datatable.GetSheet("testData").AddParameter "Test",""
datatable.getSheet("testData").SetCurrentRow(5)
datatable("Test","testData")="xxxxxxxxxxxx"


' Export complete datatable to xls file - 2003
datatable.Export "C:\qtp.xls"
datatable.ExportSheet "C:\qtp1.xls","Action1"






