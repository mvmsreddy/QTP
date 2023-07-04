closeAllBrowsers
environment.LoadFromFile "E:\Tools\QTP\Recording Scripts\DataDrivenFramework\env.xml"

openBrowser
navigate "google.com"
openBrowser
navigate "yahoo.com"
openBrowser
navigate "bbc.com"


 closeBrowser '' bbc
 setBrowserIndex(0)
 'closeBrowser '' yahoo
 closeBrowser ' google







