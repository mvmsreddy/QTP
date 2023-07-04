

environment.LoadFromFile "E:\Tools\QTP\Recording Scripts\DataDrivenFramework\env.xml"

 environment("Browser") = "Mozilla"
 openBrowser
 navigate Environment("TestSiteURL")
 defaultLogin


