'Wscript.shell

Set WshShell = createObject("Wscript.shell")
WshShell.Run "cmd"
WshShell.SendKeys "C:"
WshShell.SendKeys "~" ' enter
WshShell.SendKeys "dir"
WshShell.SendKeys "~"

