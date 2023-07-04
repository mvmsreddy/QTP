Browser("creationtime:=0").Navigate "gmail.com"
Set mypage = Browser("creationtime:=0").Page("title:=.*")

mypage.WebEdit("xpath:=//input[@id='Email']").Set "some username"
mypage.WebEdit("html id:=Email").Set "xxxx"
print mypage.WebElement("xpath:=//div[@class='product-info mail']/p").GetROProperty("innerhtml")


''' Yahoo
Browser("creationtime:=0").Navigate "yahoo.com"

Set mypage = Browser("creationtime:=0").Page("title:=.*")
mypage.webedit("class:=input-query input-long med-large  ").click
' keys simulation
Set o = createobject("wScript.shell")
o.SendKeys "hello"
wait(5)
part1="//*[starts-with(@id, 'yui_3_4_0_1_1')]/ul/li["
part2="]/a"
For i=1 to 10
print mypage.Link("xpath:="& part1 & i & part2).getroproperty("innertext")
Next









