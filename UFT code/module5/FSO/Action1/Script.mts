On error resume next
Set fso=CreateObject("Scripting.FileSystemObject")
file_Path="C:\myFile.txt"

' Create
'fso.CreateTextFile file_Path

' Write
Set myfile = fso.OpenTextFile (file_Path,2,true)
If err.number<>0 Then
	print err.number &"  - "& err.description
End If

myfile.Write "THis is the first line"
myfile.Writeline "THis is the second line"
myfile.Write "THis is the third line"

Set myfile = Nothing  ' destroy the object
'Read from a text file
Set myfile2 = fso.OpenTextFile (file_Path,1,true)
print myfile2.ReadLine
print myfile2.ReadLine

While myfile2.AtEndOfStream <> true
print myfile2.ReadLine
Wend
Set myfile2 =Nothing ' destroy
print "******************ALL Files************************"
Set f = fso.GetFolder("E:\Tools")
   Set fc = f.Files
   For Each f1 in fc
      print f1.name 
   Next
Set f=nothing  'destroy
Set fc =nothing ' destroy
print "******************ALL Folders************************"
   Set f = fso.GetFolder("E:\Tools")
   Set fc = f.SubFolders
   For Each f1 in fc
      print f1.name 
   Next




















