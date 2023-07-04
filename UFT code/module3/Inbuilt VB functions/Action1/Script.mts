' http://www.vbforums.com/showthread.php?t=316508
str="today is a rainy day and I was stuck"
msgbox left (str,5)
msgbox right(str,5)
msgbox mid(str,6)
msgbox mid(str,6,9)

str="today is a rainy day and I was stuck"
'x=split(str," ")
x=split(str,"and")
msgbox ubound(x)

For i=0 to ubound(x)
print x(i)
Next

' USCASE AND LCASE

msgbox ucase("abcdEFGH")
msgbox lcase("abcdEFGH")

' TRIM
str = "       Tom is a boy    "
print ltrim(str)
print rtrim(str)
print trim(str)


' isnumeric, cstr,cint
a="22zz"
b="44"
msgbox isnumeric(a)
If isnumeric(a) and isnumeric(b) Then
	msgbox cint(a)   + cint(b)
End If

c=1829
str=cstr(c)

'  DATE and time functions
print date
print time
print now

d1= datevalue("10 november  2012")
d2= datevalue("15 november  2012")

msgbox datediff("s",d1,d2)
msgbox datepart("d",d1)
msgbox datepart("m",d1)



'TIME
print time
print hour(time)
print minute(time)
print second(time)





















