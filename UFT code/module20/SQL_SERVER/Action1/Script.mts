Set con = CreateObject("Adodb.connection")

' connection string
connstr = "DSN=QTP_SQL_SERVER;UID=sa;PWD=pass!@#123;APP=QuickTest Professional;WSID=A-2265BA1FDCD64;DATABASE=qtp_practice_database"

con.Open connstr  ' establish connection with db


'READING DB
Set rs = con.Execute("select * from Location")

While not rs.EOF
print rs.Fields("City")  & " ----- " & rs.Fields("Country")
rs.MoveNext
Wend

' WRITE IN DB
con.Execute("insert into Location values('Tokyo','Japan')")


con.Close  ' close connection








