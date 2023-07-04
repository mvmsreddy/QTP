' Global , Local


' Interact with global sheet
print "start"
' open browser
' logging in
print datatable("Name")
datatable.SetCurrentRow(2)  ' cursor to row 2
print datatable("Name")
datatable.SetCurrentRow(3) 
print datatable("Name")
print "end"


' for loop
For i=1 to datatable.GetRowCount
datatable.SetCurrentRow(i)
print datatable("Name")
datatable("Country")="India"
Next
wait(5)

print datatable.GetRowCount



For i=1 to 10
datatable.SetCurrentRow(i)
datatable("Col1")=i
datatable("Col2")= 11-i
datatable("Col3")=cint(datatable("Col1"))+ cint(datatable("Col2"))
Next


'print cint("4") + cint("8")



datatable.GlobalSheet.AddParameter "Test_Col","ABC"
'datatable("Test_Col")="ZXXX"









