randomize
print rnd
print rnd
print rnd
print randomnumber(5,1000)

str="abcdefghijklmnopqrstuvwxyz"
' random char
' random string 4,5
print mid(str,randomnumber(1,27),1)
generateRandomString 10
generateRandomString 20

Function generateRandomString(n)
For i=1 to n
s=s+mid(str,randomnumber(1,27),1)
Next
print s
End Function

