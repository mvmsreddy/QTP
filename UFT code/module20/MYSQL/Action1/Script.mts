connStr="DSN=QTP_MYSQL_DRIVER;"
Set con =createobject("Adodb.connection")

con.Open connStr

Set rs = con.Execute("Select * from Location")


While not rs.EOF
	print rs.Fields("City") & " --- " &rs.Fields("Country")
	rs.MoveNext
Wend


con.Close






