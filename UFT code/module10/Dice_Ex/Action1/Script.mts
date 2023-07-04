'http://seeker.dice.com/jobsearch/servlet/JobSearch?op=300&N=0&Hf=0&NUM_PER_PAGE=30&Ntk=JobSearchRanking&Ntx=mode+matchall&AREA_CODES=&AC_COUNTRY=1525&QUICK=1&ZIPCODE=&RADIUS=64.37376&ZC_COUNTRY=0&COUNTRY=1525&STAT_PROV=0&METRO_AREA=33.78715899%2C-84.39164034&TRAVEL=0&TAXTERM=0&SORTSPEC=0&FRMT=0&DAYSBACK=30&LOCATION_OPTION=2&FREE_TEXT=qtp&WHERE=

i=0
Set mypage=Browser("creationtime:=0").Page("title:=.*")
While mypage.WebElement("class:=treeShow","index:="&i).exist
mypage.WebElement("class:=treeShow","index:="&i).click
wait(5)
i=i+1
Wend


i=0
Set mypage=Browser("creationtime:=0").Page("title:=.*")
While mypage.Link("class:=TC.*","index:="&i).exist
	print mypage.Link("class:=TC.*","index:="&i).getroproperty("innertext")
	mypage.Link("class:=TC.*","index:="&i).click
	mypage.sync
	print "Title ->" & mypage.GetROProperty("title")
	Browser("creationtime:=1").Close  ' close the tab
	i=i+1
Wend







