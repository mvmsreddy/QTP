datatable.AddSheet "T1"
datatable.AddSheet "T2"

datatable.ImportSheet "C:\qtp1.xls","T1","T1"
datatable.ImportSheet "C:\qtp1.xls","T2","T2"
datatable.AddSheet "Result"

compare("Name")
compare("Email")

datatable.ExportSheet "C:\qtp1.xls","Result"
Function compare(colName)
datatable.GetSheet ("Result").AddParameter colName,""
' which col has max rows
rows=0
If datatable.GetSheet("T1").GetRowCount > datatable.GetSheet("T2").GetRowCount Then
	rows = datatable.GetSheet("T1").GetRowCount
else
	rows = datatable.GetSheet("T2").GetRowCount
End If

For i=1 to rows
	datatable.GetSheet("T1").SetCurrentRow(i)
	datatable.GetSheet("T2").SetCurrentRow(i)
	datatable.GetSheet("Result").SetCurrentRow(i)
	If datatable(colName,"T1") <> datatable(colName,"T2") Then
		datatable(colName,"Result") ="N"
	else
		datatable(colName,"Result") ="Y"
	End If
Next

End Function



