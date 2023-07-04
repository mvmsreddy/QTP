environment.LoadFromFile "E:\Tools\QTP\Recording Scripts\DataDrivenFramework\env.xml"
setBrowserIndex(0)
'print verifyEnabled("GmailEdit")
'
'print verifyProperty("GmailEdit","html id","Email")

wait(5)
print isHiddenObject("ebaySell")
hovermouse("SellHeader")
print isHiddenObject("ebaySell")

