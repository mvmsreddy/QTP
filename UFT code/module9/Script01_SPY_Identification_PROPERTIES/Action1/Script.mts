'GetROProperty  - Every object - Get the runtime object property
print Browser("Quikr Classifieds: Post").Page("Quikr Classifieds: Post").WebEdit("query").GetROProperty("default value")
print Browser("Quikr Classifieds: Post").Page("Quikr Classifieds: Post").WebEdit("query").GetROProperty("width")

print Browser("Quikr Classifieds: Post").Page("Quikr Classifieds: Post").GetROProperty("title")
print Browser("Quikr Classifieds: Post").Page("Quikr Classifieds: Post").GetROProperty("url")
print Browser("Quikr Classifieds: Post").GetROProperty("application version")

print Browser("Quikr Classifieds: Post").Exist
print Browser("Quikr Classifieds: Post").Page("Quikr Classifieds: Post").Exist
print Browser("Quikr Classifieds: Post").Page("Quikr Classifieds: Post").WebEdit("query").Exist

If  Browser("Quikr Classifieds: Post").Page("Quikr Classifieds: Post").WebElement("- Free local classifieds").Exist Then
	print Browser("Quikr Classifieds: Post").Page("Quikr Classifieds: Post").WebElement("- Free local classifieds").GetROProperty("innertext")
End If



