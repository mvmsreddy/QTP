' parametrizing checkpoint
For i=1 to datatable.GetRowCount
datatable.SetCurrentRow(i)
domain = datatable("Domains")
Browser("Mozilla Firefox Start").DeleteCookies
Browser("Mozilla Firefox Start").Page("Gmail: Email from Google").Sync @@ hightlight id_;_Browser("Mozilla Firefox Start").Page("Gmail: Email from Google")_;_script infofile_;_ZIP::ssf3.xml_;_
Browser("Mozilla Firefox Start").navigate "fatcow.com"

Browser("Mozilla Firefox Start").Page("Web Hosting & Domain Names").Link("Get Started!").Click @@ hightlight id_;_Browser("Mozilla Firefox Start").Page("Web Hosting & Domain Names").Link("Get Started!")_;_script infofile_;_ZIP::ssf7.xml_;_
Browser("Mozilla Firefox Start").Page("Web Hosting by FatCow").WebElement("WebElement").Click

ret = Browser("Mozilla Firefox Start").Page("FatCow Registration").WebEdit("dom_lookup").Check (CheckPoint("dom_lookup")) @@ hightlight id_;_Browser("Mozilla Firefox Start").Page("FatCow Registration").WebEdit("dom lookup")_1_;_script infofile_;_ZIP::ssf4.xml_;_
If not ret Then
	exitaction "field was not empty"
End If
Browser("Mozilla Firefox Start").Page("FatCow Registration").WebEdit("dom_lookup").Set domain @@ hightlight id_;_Browser("Mozilla Firefox Start").Page("FatCow Registration").WebEdit("dom lookup")_;_script infofile_;_ZIP::ssf4.xml_;_
Browser("Mozilla Firefox Start").Page("FatCow Registration").WebButton("Continue").Click @@ hightlight id_;_Browser("Mozilla Firefox Start").Page("FatCow Registration").WebButton("Continue")_;_script infofile_;_ZIP::ssf5.xml_;_

Browser("Mozilla Firefox Start").Page("FatCow Registration").Check CheckPoint("FatCow Registration") @@ hightlight id_;_Browser("Mozilla Firefox Start").Page("FatCow Registration")_;_script infofile_;_ZIP::ssf6.xml_;_
Browser("Mozilla Firefox Start").Page("FatCow Registration").WebEdit("FirstName").Set "xxxx" @@ hightlight id_;_Browser("Mozilla Firefox Start").Page("FatCow Registration").WebEdit("FirstName")_;_script infofile_;_ZIP::ssf6.xml_;_
next
