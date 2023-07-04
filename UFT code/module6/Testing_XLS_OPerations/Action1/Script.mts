msgbox readData("c:\xls\testApp.xls","test1",2,2) ' read the data
writeData  "c:\xls\testApp.xls","test1",5,5,"myTestData"


writeLink "c:\xls\testApp.xlsx","test1",5,5,"myTestData","C:\xls\temp.png"


' total data rows in sheet
'total cols in the sheet

' write these functions - Your work
' if the col is existing
' delete the exising col

print getColumnCount ("c:\xls\data.xlsx","Login")

print "Total rows - " & getRowCount ("c:\xls\data.xlsx","Login")



writeData  "c:\xls\data.xlsx","test1",5,5,"myTestData"


