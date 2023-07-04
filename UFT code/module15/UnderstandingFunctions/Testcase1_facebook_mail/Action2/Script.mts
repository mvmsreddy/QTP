 Browser("Rediff.com - India, Business,").Page("Rediffmail NG").WebEdit("searchInput").Set "ss"


'' open latest mail from facebook
' Without OR


Browser("creationtime:=0").page("title:=.*").WebElement("innerhtml:=Facebook","index:=0").Click