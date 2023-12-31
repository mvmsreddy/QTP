 Option explicit
 Dim selectedBrowserIndex
selectedBrowserIndex=-1


''' Opens a new Browser
Function openBrowser
   
   If environment("Browser") = "Mozilla" and NOT isBrowserOpen("Mozilla") then
		SystemUtil.Run "H:\Program Files\Mozilla Firefox\firefox.exe"
		setBrowserIndex getBrowserIndex+1
   elseif environment("Browser") = "IE" and NOT isBrowserOpen("internet explorer") then
        SystemUtil.Run "H:\Program Files\Internet Explorer\IEXPLORE.EXE"
		setBrowserIndex getBrowserIndex+1
	elseif environment("Browser") = "Mozilla" and  isBrowserOpen("Mozilla") then
	' bring the control on mozilla
		setBrowserIndex(0) '' making sure mozilla will have 0 index

	elseif environment("Browser") = "IE" and  isBrowserOpen("internet explorer") then
	' bring the control on IE
		setBrowserIndex(1) '' making sure mozilla will have 0 index

   end if
	If  selectedBrowserIndex = -1 Then
		setBrowserIndex(0)
	End If

sync ' let the page in browser open

End Function


'''' Sets the browser Index
Function setBrowserIndex(bIndex)
selectedBrowserIndex = bIndex
End Function

'' Get the browser Index
Function getBrowserIndex
getBrowserIndex = selectedBrowserIndex
End Function

' Returns the Page object
Function getPage
   Set getPage=Browser("creationtime:="&getBrowserIndex).Page("title:=.*")
End Function

'Close the current browser
Function closeBrowser
   Browser("creationtime:="&getBrowserIndex).close
   setBrowserIndex getBrowserIndex-1
End Function

' close all browsers
Function closeAllBrowsers
   Dim desc,objBrowsers,bNum
   Set desc = description.Create
   desc("micclass").value="Browser"
   Set objBrowsers = desktop.ChildObjects(desc)

	For bNum=0 to objBrowsers.count-1
		objBrowsers(bNum).close
	Next
	setBrowserIndex -1

	' destroy objects
	Set desc=nothing
	Set objBrowsers=nothing
End Function

' page sync
Function sync
   getPage().sync ' let the current page load
End Function

' Navigate the current browser to a URL
Function navigate(URL)
   Browser("creationtime:="&getBrowserIndex).navigate(URL)
   sync
End Function

'''''''''''''''''''''''''Whether browser is opened or not''''''''''''''''''''''''''''
Function isBrowserOpen(bName)

Set desc_browser= description.Create
desc_browser("micclass").value="Browser"

Set objBrowser = desktop.ChildObjects(desc_browser)
For i=0 to objBrowser.count-1
 If instr(objBrowser(i).getROProperty("application version"),bName) > 0 then
	 isBrowserOpen=true
	 Set objBrowser=nothing
	 Exit function
 end if
Next
isBrowserOpen=false
Set objBrowser=nothing
End Function


