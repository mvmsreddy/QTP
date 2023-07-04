'Print 
msgbox "hello world"

'Integer, String, decimal, boolean

a=100
b=200
msgbox a
msgbox b
msgbox a+b

a="We are learning"
msgbox a


c=49.45
c=16.27
print c
print " qtp is a wise tool "

'  InputBox
n = inputbox ("Enter a number","Number application")
print n
' Concatination operator
print " Value entered by the user is " & n


x=100
y=200
print x & y

'  Big strings
mailMessage="hello david.i am fine. How are you ?" &_
"I am fine. Howzz it going" &_
"Regards"
msgbox mailMessage

' Breaking line
temp = "There is a cat." & vbcrlf  &  "There is a bird."
msgbox temp

mailMessage="hello david.i am fine. How are you ?" & vbcrlf &_
"I am fine. Howzz it going"  & vbcrlf &_
"Regards"
msgbox mailMessage


