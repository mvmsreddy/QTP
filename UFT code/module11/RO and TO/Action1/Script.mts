



expected_text = Browser("Monster - Jobs in India,Delhi,").Page("Monster - Jobs in India,Delhi,").Link("Banking/ Financial Services").GetTOProperty("text")
actual_text = Browser("Monster - Jobs in India,Delhi,").Page("Monster - Jobs in India,Delhi,").Link("Banking/ Financial Services").GetROProperty("text")
print expected_text
print actual_text


Set mypage = Browser("Monster - Jobs in India,Delhi,").Page("Monster - Jobs in India,Delhi,")
Set allProp = mypage.Link("Banking/ Financial Services").GetTOProperties
print allProp.count

For i=0 to allProp.count-1
print allProp(i).name &"  ----  "& allProp(i).value  ' OR
print mypage.Link("Banking/ Financial Services").GetROProperty(allProp(i).name)
print "------------------"
expected=allProp(i).value 
actual=mypage.Link("Banking/ Financial Services").GetROProperty(allProp(i).name)

If expected<>actual Then
	reporter.ReportEvent micFail,"monster","Found "& actual
End If
Next







