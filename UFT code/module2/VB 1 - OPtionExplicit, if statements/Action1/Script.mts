Option explicit
Dim a
a=100
msgbox a

a="In the world of qtp"
msgbox a

' If statements
msgbox 3<2
x=true ' true or false

' Ctr + F5
a="Hello1"
b="Hello"

x=200
y=200
'  >   <  =  <>  >=   <=
If x>y Then
	msgbox "X is greatest "& x
elseif x<y then
	msgbox "Y is greatest "& y
else
    msgbox "Both are equal"
End If


x=100
y=200
z=50

If x>y and x>z Then '' boolean operator - and or xor
	msgbox "X is greatest "& x
elseif y>z then
	msgbox "Y is greatest "& y
else
   msgbox "z is greatest "& z
End If















