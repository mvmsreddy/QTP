' Google time
str="10:12 Tuesday (BST) - Time in London, UK"
pat="[0-9]{1,2}:[0-9]{1,2} (Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday) [(]BST[)] - Time in London, UK"

findPatternMatches str,pat

print Browser("time in london - Google").Page("time in london - Google").WebElement("TimeandDay").GetROProperty("innertext")


''  BBC
text_description =  Browser("time in london - Google").Page("BBC - Homepage").WebElement("UN-Arab League envoy Kofi").GetROProperty("innertext")
pat="[0-9a-zA-Z].*"
res = findPatternMatches (text_description,pat)

If res Then
	print "found"
	'reporter
else 	
	print "not found"
	'reporter
	'screenshot
End If

'Webtable
Set obj_table = Browser("time in london - Google").Page("Gaana.com Listen Songs").WebTable("Pani Da Rang (...")
print obj_table.RowCount

Dim x
'For rowNum=1 to obj_table.RowCount
'x=""
'	For colNum=1 to obj_table.ColumnCount(rowNum)
'		x = x & obj_table.GetCellData(rowNum,colNum) & " - "
'	Next
'print x
'Next

For rowNum=1 to obj_table.RowCount
 str= obj_table.GetCellData(rowNum,1)
 pat="[0-9a-zA-Z].*"
 print findPatternMatches (str,pat)
next

' http://money.rediff.com
Dim amt
amt = Browser("time in london - Google").Page("Ambuja Cements Ltd. Stock").WebElement("ProfitLoss").GetROProperty("innertext")
print amt
pat="[0-9].*[.][0-9]{2}"
print findPatternMatches (amt,pat)




