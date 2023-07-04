'  xls   xlsx

'worbooks   - xls file
' sheet - cheet in xls file, 
option explicit 
Dim objXls, myXls, sheet1, sheet2, exitingSheeting, filePath,fso
filePath="C:\xls\testApp.xlsx"
Set objXls = createobject("Excel.application")  '' microsoft  xl 
objXls.Visible= true
' created an xls file
Set fso = createobject("Scripting.FileSystemObject")

If NOT fso.FileExists(filePath) Then
   Set myXls = objXls.Workbooks.Add
else
   'Set myXls= objXls.workbooks.Open (filePath)
   ' delete the file and create a new workbook
	fso.DeleteFile filePath
	   Set myXls = objXls.Workbooks.Add
End If



' add sheet in xls
Set sheet1 = myXls.Sheets.add
sheet1.name="test1"

Set sheet2 = myXls.Sheets.add
sheet2.name="test2"

' delete test2 sheet
sheet2.delete

' deleting existing sheet
Set exitingSheeting = myXls.Sheets("sheet1")
exitingSheeting.delete

sheet1.cells(2,3).value = "city"

sheet1.cells(5,5).value  = "=Hyperlink("&chr(34)&"C:\xls\temp.png"&chr(34)&", "&chr(34)&"click"&chr(34)&")"

msgbox sheet1.cells(2,3).value


objXls.DisplayAlerts = false
myXls.SaveAs "C:\xls\testApp.xls"   ' office 2003
myXls.SaveAs "C:\xls\testApp.xlsx"   ' office 2007

myXls.Close '  close the xl file
objXls.Quit


msgbox chr(34)









