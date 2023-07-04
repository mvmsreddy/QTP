 print Browser("Quikr Classifieds: Post").Page("Quikr Classifieds: Post").WebEdit("query").GetROProperty("value")
Browser("Quikr Classifieds: Post").Page("Quikr Classifieds: Post").WebList("categoryId").Select "Chandigarh" @@ hightlight id_;_Browser("Quikr Classifieds: Post").Page("Quikr Classifieds: Post").WebList("categoryId")_;_script infofile_;_ZIP::ssf1.xml_;_


Browser("Quikr Classifieds: Post").Page("Quikr Classifieds: Post").WebEdit("query").Output CheckPoint("query") @@ hightlight id_;_Browser("Quikr Classifieds: Post").Page("Quikr Classifieds: Post").WebEdit("query")_;_script infofile_;_ZIP::ssf1.xml_;_
print datatable("query_value_out") @@ hightlight id_;_Browser("Quikr Classifieds: Post").Page("Quikr Classifieds: Post").Link("Akola")_;_script infofile_;_ZIP::ssf2.xml_;_



Browser("Quikr Classifieds: Post").Page("Quikr Classifieds: Post").WebElement("- Free local classifieds for apartments, cars, jobs, services, used goods, events and more.").Output CheckPoint("quikr welcome text") @@ hightlight id_;_Browser("Quikr Classifieds: Post").Page("Quikr Classifieds: Post").WebElement("- Free local classifieds for apartments, cars, jobs, services, used goods, events and more.")_;_script infofile_;_ZIP::ssf1.xml_;_
print datatable("mycol")


XMLFile("student.xml").Output CheckPoint("student.xml")
print datatable("value_name_out")
