 If  Window("Calculator").WinButton("5").Exist Then
	 print Window("Calculator").WinButton("5").GetROProperty("text")
 End If

 Window("Calculator").WinButton("3").Click

print  Window("Calculator").WinEdit("Edit").GetROProperty("text")
 Window("Calculator").WinButton("*").Click
 Window("Calculator").WinButton("8").Click

print  Window("Calculator").WinEdit("Edit").GetROProperty("text")
Window("Calculator").WinButton("=").Click
print  Window("Calculator").WinEdit("Edit").GetROProperty("text")




s= Browser("The Times of India: Latest").Page("The Times of India: Latest").Frame("Frame").WebElement("total").GetROProperty("innertext")
temp = split(s," ")
print temp(0)
print temp(1)
print temp(2)
print temp(3)

' number from app - string
sum = cint(temp(0)) + cint(temp(2))
print "Sum is -- " & sum
Browser("The Times of India: Latest").Page("The Times of India: Latest").Frame("Frame").WebEdit("mathuserans2").Set sum

print 5+3 ' 8
print "5" + "3" ' 53
print cint("5") + cint("3") ' 8
'cstr

Dim oldtext, newtext

oldtext = Browser("The Times of India: Latest").Page("Gmail: Email from Google").WebElement("Increasing number").GetROProperty("innertext")
i=1
While  i<>6
	newtext = Browser("The Times of India: Latest").Page("Gmail: Email from Google").WebElement("Increasing number").GetROProperty("innertext")
    If  oldtext <> newtext Then
		print newtext
		oldtext = newtext
		i=i+1
	End If
wend






