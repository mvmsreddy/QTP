 print Browser("Quikr Classifieds: Post").Page("Quikr Classifieds: Post").WebButton("Search").GetROProperty("value")



' quikr.com
 Browser("Quikr Classifieds: Post").Page("Quikr Classifieds: Post").WebList("categoryId").Select "Goa"  ' starts from 0

 Browser("Quikr Classifieds: Post").Page("Quikr Classifieds: Post").WebList("categoryId").Select 4

totalEle = Browser("Quikr Classifieds: Post").Page("Quikr Classifieds: Post").WebList("categoryId").GetROProperty("items count")
print "Total Elements in the list are " & totalEle
print Browser("Quikr Classifieds: Post").Page("Quikr Classifieds: Post").WebList("categoryId").GetItem(1)   '  starts from 1

print "All elements"
For i=1 to totalEle
	print Browser("Quikr Classifieds: Post").Page("Quikr Classifieds: Post").WebList("categoryId").GetItem(i)   '  starts from 1
Next

currSelectedVal = Browser("Quikr Classifieds: Post").Page("Quikr Classifieds: Post").WebList("categoryId").GetROProperty("value")
print "Current Selected value--- " & currSelectedVal


' Multi select List

Browser("Quikr Classifieds: Post").Page("Multi-select Checkboxes").WebList("options").Select "Option 1" @@ hightlight id_;_Browser("Quikr Classifieds: Post").Page("Multi-select Checkboxes").WebList("options")_;_script infofile_;_ZIP::ssf1.xml_;_
Browser("Quikr Classifieds: Post").Page("Multi-select Checkboxes").WebList("options").ExtendSelect "Option 2" @@ hightlight id_;_Browser("Quikr Classifieds: Post").Page("Multi-select Checkboxes").WebList("options")_;_script infofile_;_ZIP::ssf2.xml_;_
Browser("Quikr Classifieds: Post").Page("Multi-select Checkboxes").WebList("options").ExtendSelect "Option 4" @@ hightlight id_;_Browser("Quikr Classifieds: Post").Page("Multi-select Checkboxes").WebList("options")_;_script infofile_;_ZIP::ssf3.xml_;_


Browser("Quikr Classifieds: Post").Page("Multi-select Checkboxes").WebList("options").Select 1 @@ hightlight id_;_Browser("Quikr Classifieds: Post").Page("Multi-select Checkboxes").WebList("options")_;_script infofile_;_ZIP::ssf1.xml_;_
Browser("Quikr Classifieds: Post").Page("Multi-select Checkboxes").WebList("options").ExtendSelect 3 @@ hightlight id_;_Browser("Quikr Classifieds: Post").Page("Multi-select Checkboxes").WebList("options")_;_script infofile_;_ZIP::ssf2.xml_;_
Browser("Quikr Classifieds: Post").Page("Multi-select Checkboxes").WebList("options").ExtendSelect 4 @@ hightlight id_;_Browser("Quikr Classifieds: Post").Page("Multi-select Checkboxes").WebList("options")_;_script infofile_;_ZIP::ssf3.xml_;_