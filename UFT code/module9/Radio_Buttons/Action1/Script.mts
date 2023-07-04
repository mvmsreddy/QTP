Browser("IndiGo - On Time, Low").Page("IndiGo - On Time, Low").WebRadioGroup("radioSector1").Select "#3"  ' starts from 0 @@ hightlight id_;_Browser("IndiGo - On Time, Low").Page("IndiGo - On Time, Low").WebRadioGroup("radioSector1")_;_script infofile_;_ZIP::ssf1.xml_;_
print Browser("IndiGo - On Time, Low").Page("IndiGo - On Time, Low").WebRadioGroup("radioSector1").GetROProperty("items count")
print Browser("IndiGo - On Time, Low").Page("IndiGo - On Time, Low").WebRadioGroup("radioSector1").GetROProperty("selected item index")



