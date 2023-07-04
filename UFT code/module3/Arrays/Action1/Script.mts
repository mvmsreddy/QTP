
Dim x(5)  ' 6 elements
x(0)=10
x(1)=18
x(2)="Hello world"
x(3)=true
x(4)=181.43
x(5)=19
print "The size of array is " & ubound(x)
print x(2)

For i=0 to ubound(x)
print x(i)
Next

For i=ubound(x) to 0 step -1
print x(i)
Next

'' Redimensional arrays
Dim y()
ReDim preserve y(2) ' 3
y(0)=10
y(1)=12
y(2)=13

ReDim preserve y(4) '5
y(3)=18
y(4)=19

For i=0 to ubound(y)
print y(i)
Next

