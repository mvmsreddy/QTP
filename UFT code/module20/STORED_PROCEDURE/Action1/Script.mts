Set con = createobject("Adodb.Command")
con.ActiveConnection = "DSN=QTP_SQL_SERVER;UID=sa;PWD=pass!@#123;APP=QuickTest Professional;WSID=A-2265BA1FDCD64;DATABASE=qtp_practice_database"


con.CommandType=4 '  inform that you are calling a stored procdure
con.CommandText="locations_world"  ' name of the stored procedure

con.Parameters(1).Value="Sydney"  ' 1st input parameter
con.Parameters(2).Value=""  '' last ouput parameter

con.Execute

print con.Parameters(0).Value ' result of executing sp
print con.Parameters(2).Value


