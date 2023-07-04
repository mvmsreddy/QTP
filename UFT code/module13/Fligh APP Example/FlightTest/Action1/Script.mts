environment.LoadFromFile"E:\Tools\QTP\Recording Scripts\module 15\Fligh APP Example\environment.xml"
datatable.AddSheet "Fight Data"
datatable.ImportSheet environment("DataXL"),"Reservation Details","Fight Data"


SystemUtil.Run environment("ApplicationLocation")
Dialog("Login").WinEdit("Agent Name:").Set environment("LoginUserID") @@ hightlight id_;_526084_;_script infofile_;_ZIP::ssf1.xml_;_
Dialog("Login").WinEdit("Agent Name:").Type  micTab  @@ hightlight id_;_526084_;_script infofile_;_ZIP::ssf2.xml_;_
Dialog("Login").WinEdit("Password:").SetSecure environment("Password") @@ hightlight id_;_919480_;_script infofile_;_ZIP::ssf3.xml_;_
Dialog("Login").WinButton("OK").Click @@ hightlight id_;_788476_;_script infofile_;_ZIP::ssf4.xml_;_
wait(5)

For j=1 to datatable.GetSheet("Fight Data").GetRowCount
datatable.GetSheet("Fight Data").SetCurrentRow(j)
Window("Flight Reservation").ActiveX("MaskEdBox").Type datatable("Date","Fight Data") @@ hightlight id_;_788456_;_script infofile_;_ZIP::ssf5.xml_;_
Window("Flight Reservation").WinComboBox("Fly From:").Select datatable("Source","Fight Data") @@ hightlight id_;_591774_;_script infofile_;_ZIP::ssf6.xml_;_
Window("Flight Reservation").WinComboBox("Fly To:").Select datatable("Destination","Fight Data") @@ hightlight id_;_591778_;_script infofile_;_ZIP::ssf7.xml_;_
Window("Flight Reservation").WinButton("FLIGHT").Click @@ hightlight id_;_854012_;_script infofile_;_ZIP::ssf8.xml_;_
selectMinPricedFlight()
'Window("Flight Reservation").Dialog("Flights Table").WinList("From").Select "6284   DEN   12:48 PM   LAX   03:48 PM   AA     $102.50" @@ hightlight id_;_198734_;_script infofile_;_ZIP::ssf12.xml_;_
Window("Flight Reservation").Dialog("Flights Table").WinButton("OK").Click @@ hightlight id_;_198742_;_script infofile_;_ZIP::ssf13.xml_;_
Window("Flight Reservation").WinEdit("Name:").Set datatable("Name","Fight Data") @@ hightlight id_;_985016_;_script infofile_;_ZIP::ssf14.xml_;_
Window("Flight Reservation").WinRadioButton("First").Set @@ hightlight id_;_591620_;_script infofile_;_ZIP::ssf15.xml_;_
Window("Flight Reservation").WinButton("Insert Order").Click @@ hightlight id_;_526114_;_script infofile_;_ZIP::ssf16.xml_;_
' pull out the order number
While Window("Flight Reservation").WinEdit("Order No:").GetROProperty("text")=""
Wend
print Window("Flight Reservation").WinEdit("Order No:").GetROProperty("text")
datatable("OrderNumber","Fight Data")=Window("Flight Reservation").WinEdit("Order No:").GetROProperty("text")
Window("Flight Reservation").WinMenu("Menu").Select "File;New Order"


Next

datatable.ExportSheet environment("FinalXL"),"Fight Data"

Function selectMinPricedFlight()

Set flightList = Window("Flight Reservation").Dialog("Flights Table").WinList("From")
'msgbox flightList.GetContent
print  flightList.GetItem(0)
print flightList.GetROProperty("items count")

print "*****************ALL ITEMS***********************"

min_price=0
min_price_index=0
For i =0 to flightList.GetROProperty("items count")-1
print flightList.GetItem(i)
arr = split(flightList.GetItem(i),"$")

If i=0 Then
	min_price=arr(1)
	min_price_index=i
elseif min_price > arr(1) then
	min_price=arr(1)
	min_price_index=i
End If
Next

print "Min price is " & min_price

flightList.Select min_price_index


End Function






















