'writeData "E:\Tools\QTP\Recording Scripts\DataDrivenFramework\Testing the scripts-Rough\Test.xlsx","Countries",2,4,"QTP"

print getColumnCount ("E:\Tools\QTP\Recording Scripts\DataDrivenFramework\Testing the scripts-Rough\Test.xlsx","Countries")

print IsColumnExisting("E:\Tools\QTP\Recording Scripts\DataDrivenFramework\Testing the scripts-Rough\Test.xlsx","Countries","States")
print IsColumnExisting("E:\Tools\QTP\Recording Scripts\DataDrivenFramework\Testing the scripts-Rough\Test.xlsx","Countries","Hello")



print getRowCount ("E:\Tools\QTP\Recording Scripts\DataDrivenFramework\Testing the scripts-Rough\Test.xlsx","Countries")
print getRowCount ("E:\Tools\QTP\Recording Scripts\DataDrivenFramework\Components.xlsx","App Independent Libs")
destroyFile


addSheet "E:\Tools\QTP\Recording Scripts\DataDrivenFramework\Components.xlsx","TestSheet"
compareSheets "E:\Tools\QTP\Recording Scripts\DataDrivenFramework\Testing the scripts-Rough\Test.xlsx","A","B","Res"



Set x = getData ("E:\Tools\QTP\Recording Scripts\DataDrivenFramework\Testing the scripts-Rough\Test.xlsx","Data","BookTicket")
print x.item(1).item("State")
print x.item(2).item("State")
destroyFile







