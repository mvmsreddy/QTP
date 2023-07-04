str="this day this week this month this year"
pat="this"
findPatternMatches str,pat


Set regex = new RegExp
regex.Global =true
regex.ignorecase=true
regex.pattern=pat

print "Pattern Found - " & regex.test(str)
' index, string

' matches collection and match object
Set matches = regex.execute(str)  ' collection
print matches.count

Set match = matches(0) ' object
print match.firstIndex
print match.value


Set match = matches(1) ' object
print match.firstIndex
print match.value


' for loop to extract all matches from matches collection

For i=0 to matches.count-1
Set match = matches(i)
print "Value found at index "& match.firstIndex &" . Match value is - "& match.value
Next


findPatternMatches " xxxxx yyyyy zzzz dddd","tttttt"

Function findPatternMatches(str,pat)

Set regex = new RegExp
regex.Global =true
regex.ignorecase=true
regex.pattern=pat

print "Pattern Found - " & regex.test(str)

' matches collection and match object
Set matches = regex.execute(str)  ' collection
print "Total matches found: "& matches.count
For i=0 to matches.count-1
Set match = matches(i)
print "Value found at index "& match.firstIndex &" . Match value is - "& match.value
Next

End Function





