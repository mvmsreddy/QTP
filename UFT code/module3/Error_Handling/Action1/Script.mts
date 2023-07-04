On error resume next
print "A " & Err.number
cint("a3")
' Report error
print "B "& Err.number

If  Err.number <> 0 Then
	print Err.description
End If

Err.clear
print "After clearing - " & Err.Number


On error resume next
Err.raise 10
print err.description










