s="then there that this the abcthxyz"
p="th"
p="th.*"
p="th{1,}"
p="th+"
p="\bth+"
findPatternMatches s,p


s="510 982 392"
p="[0-9]"
p="[0-9]{1,}"
p="[0-9]{3}"
p="[0-9]{3}\b"

findPatternMatches s,p

s="The completion of selection process."
p="\b[a-z]+tion\b"
p="completion|selection"
p="(comple|selec)tion"
findPatternMatches s,p

s="Welcome Sam. You have 8282 unread mails."
p="[a-zA-Z0-9]+"
p="^Welcome [a-zA-Z]+[.]\sYou have [0-9]+ unread mails[.]"
p="([a-zA-Z]+)|([.])|([0-9]+)"
findPatternMatches s,p

s="HCQ9D-TVCWX-X9QR1-J4B2Y-2GR2T"
p="[A-Z0-9]{5}-{0,1}"
p="[A-Z]{5}-{0,1}"
findPatternMatches s,p

s="Email id is itsthakur@gmail.com or qtpselenium@gmail.com or itsthakur@yahoo.com"
p="[a-zA-Z0-9]{3,}@[a-zA-Z]{1,}[.][a-zA-Z]{1,}"
findPatternMatches s,p

s="dates are 21/01/2014  11/04/2014   21/10/2014   31/11/2014"
p="((0[1-9])|(1[0-9])|(2[0-9])|(3[0-1]))/((0[1-9])|(1[0-2]))/[0-9]{4}"
findPatternMatches s,p



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























