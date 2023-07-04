environment.LoadFromFile "E:\Tools\QTP\Recording Scripts\module 15\goindigo\environment.xml"
datatable.AddSheet "MyFlightData"
datatable.ImportSheet environment("XLFileLocation"),"Flight","MyFlightData"

For i=1 to datatable.GetSheet("MyFlightData").GetRowCount
datatable.GetSheet("MyFlightData").SetCurrentRow(i)
' book ticket
Browser("creationtime:=0").Navigate environment("TestSite")

Set myPage = Browser("creationtime:=0").Page("title:=.*")
myPage.WebRadioGroup("html id:=roundTripRadio").select "#1"

Browser("creationtime:=0").Dialog("text:=The page at.*").Page("title:=.*").WebButton("name:=OK").Click

myPage.WebList("name:=from1").Select datatable("From","MyFlightData")
myPage.WebList("name:=to1").Select datatable("To","MyFlightData")
myPage.WebList("name:=departDay1").Select datatable("Day","MyFlightData")
myPage.WebList("name:=departMonth1").Select cint(datatable("Month","MyFlightData"))
myPage.Image("file name:=b_search_home.gif").Click


Set table=myPage.WebTable("class:=outbound")
print "Total rows in table -->" &  table.RowCount
'print table.GetCellData(1,4)
max_price=-1
max_index=0
For rowNum=2 to table.RowCount
price_string =  table.GetCellData(rowNum,4)
temp = split(price_string,"Rs")
price = trim(temp(1))
If  price > max_price Then
	max_price=price
	max_index=rowNum-2
End If
Next

print "Max price is -> "& max_price
print "Max index is "& max_index
myPage.WebRadioGroup("name:=radioSector1").Select "#"&max_index

Next











