' admin username/pass
' URL - Prod, UAT
' Paths -  xls , screenshots


' Environment var - System and user
msgbox environment("OS")
msgbox environment("UserName")
' User defined env

msgbox environment("password")


environment.LoadFromFile "H:\environment.xml"

msgbox environment("admin")
msgbox environment("password")
msgbox environment("URL")
'msgbox environment("XXXXXXXXXX")

msgbox environment.ExternalFileName





