'  Calling the function
x= sum (12,13)
msgbox x
sum 4,5
sum -3,2

Function sum(a,b)  ' Declaring a function
   sum =  a+b
End Function


calSum 10
calSum 20
calSum 30


Function calSum(n)
sum=0
For i=0 to n
sum=sum+i
' sum  i
'0        0
' 0       1
'1        2
'3        3
Next

msgbox  "Sum of first "& n &" numbers is "& sum
End Function


m= mult (3,4)
msgbox "Result is "& m
If result >10 Then
	print "xxxxxx"
End If

Function mult(x,y)
   result=x*y
   mult=result
End Function

'  variable with same name as function name
' put the brackets










