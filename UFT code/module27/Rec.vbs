 
 
Function RecoveryFunction1(Object, Method, Arguments, retVal)
 print "Error in App"
 print Object.getTOproperty("name")
 print method

For i=0 to ubound(Arguments)
	print Arguments(i)
Next

 
 print retval
 print describeresult(retval)
 ' capture screenshot
End Function 
 