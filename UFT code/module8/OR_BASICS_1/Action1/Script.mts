SystemUtil.Run "H:\Program Files\Mozilla Firefox\firefox.exe","","H:\Program Files\Mozilla Firefox","open"
Browser("Mozilla").Page("Google Page").WebEdit("Search Box").Set "QTP"
Browser("Mozilla").Page("Google Page").WebButton("Google Search").Click
wait(10) '  forcing qtp to pause for 10 seconds
Browser("Mozilla").Page("Google Search Result Page").Sync ' let the full page load
Browser("Mozilla").Page("Google Search Result Page").Link("Learn Quick Test Professional").Click







