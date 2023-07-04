' count
c=0
While Browser("creationtime:="&c).exist
	c=c+1
Wend

print "total browsers - " & c

'close
While Browser("creationtime:=0").Exist
Browser("creationtime:=0").Close
Wend






