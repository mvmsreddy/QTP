print micPass  '0
print micFail   '1
print micWarning  '3
print micDone '4

Reporter.ReportNote " test case started "

SystemUtil.Run "H:\Program Files\Mozilla Firefox\firefox.exe","","H:\Program Files\Mozilla Firefox","open"
Set mypage = Browser("creationtime:=0").page("title:=.*")
Browser("creationtime:=0").Navigate "http://qtpselenium.com"
mypage.image("file name:=view_details.png").Click
mypage.Sync
'wait(5)
expected="Online Selenium Training | Free Selenium Tutorial | Learn Selenium"
actual=mypage.GetROProperty("title")
print actual
print expected
If actual <> expected Then
	' screenshot
desktop.CaptureBitmap "C:/qtpselenium.png",1
reporter.ReportEvent  micFail,"Selenium Training Page"," Title did not match.Found "& actual,"C:/qtpselenium.png"
else
reporter.ReportEvent micPass,"Selenium Training Page","Titles match"
End If


Reporter.ReportNote " test case end "

print reporter.ReportPath ' report path

If reporter.RunStatus = 1 then
	' do not continue
	'exit
end if



Reporter.ReportEvent micDone, "1", "" 
Reporter.ReportEvent micDone, "2", "" 
Reporter.Filter = rfDisableAll 

Reporter.ReportEvent micDone, "3", "" 
Reporter.ReportEvent micDone, "4", "" 

Reporter.Filter = rfEnableAll 
Reporter.ReportEvent micDone, "5", "" 
Reporter.ReportEvent micDone, "6", "" 













