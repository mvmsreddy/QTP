Set desc = description.Create
desc("micclass").value="Browser"

Set openBrowsers = desktop.ChildObjects(desc)
print "Total browsers Open ->" & openBrowsers.count

For i=0 to openBrowsers.count-1
print openBrowsers(i).getROProperty("application version")
openBrowsers(i).close
Next

openBrowsers(0).navigate "google.com"



' windows apps
Set desc = description.Create
desc("micclass").value="Dialog"

Set app = desktop.ChildObjects(desc)
print app.count

desc("micclass").value="WinEdit"
Set editB = app(0).childObjects(desc)
print editB.count
editB(0).set "xxxxx"
editB(1).set "pppp"


