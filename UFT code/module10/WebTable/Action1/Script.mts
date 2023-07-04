Set myTable= Browser("creationtime:=0").Page("title:=.*").webtable("name:=spacer")
print myTable.RowCount
print mytable.ColumnCount(1)

print mytable.GetCellData(1,1)
print "***************"
text_table=""
For rowNum=1 to myTable.RowCount

	For colNum=1 to mytable.ColumnCount(rowNum)
           text_table = text_table & myTable.GetCellData(rowNum,colNum) &" - "
	Next
	  text_table= text_table & vbcrlf

Next

print text_table
' only from a particular company
' pagination


' open erp

Set myTable= Browser("creationtime:=0").Page("title:=.*").webtable("name:=pencil","column names:=.*Order Reference.*")
print myTable.RowCount
For rownum=1 to myTable.RowCount
	print myTable.GetCellData(rownum,8)
Next
' find the total
' get rid of the error	

print myTable.GetRowWithCellText("Demo User")










