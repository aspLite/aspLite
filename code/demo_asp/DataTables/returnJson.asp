<%
dim returnJson : set returnJson=new cls_dt_returnJson

class cls_dt_returnJson

	private OrderCol,OrderDir,strOrder,draw,StartRecord,RowsPerPage,JsonAnswer,JsonHeader
	public strSearch, strWhere, dbPath, strSelect, db

	Private Sub Class_Initialize()
	
		on error resume next
	
		'connect to the database - in this case an access database (dbpath)
		set db=aspL.plugin("database")
		db.path="db/sample.mdb"
	
		'which column will have to be sorted on ?
		OrderCol = aspL.convertNmbr(aspL.getRequest("Order[0][column]"))
		
		'asc or desc?
		OrderDir =  trim(lcase(aspL.getRequest("Order[0][dir]")))
		'prevent SQLinjection
		if OrderDir<>"asc" and OrderDir<>"desc" then OrderDir="asc"
		 
		'WHERE clause uses columns number, like e.g: ORDER BY 1 DESC, you may add translations to column names here, like e.g.: OrderCol = Replace(OrderCol,"0","Col1")
		'We are adding 1 here, because DataTables indexes columns starting from 0
		strOrder=" Order By " & OrderCol+1 & " " & OrderDir
		 
		draw = aspL.convertNmbr(aspL.getRequest("draw"))
		
		'where exactly are we in the table?
		StartRecord = aspL.convertNmbr(aspL.getRequest("start"))
		
		if StartRecord = 0 then 
			StartRecord=1 'Absolutepage cannot be 0
		else
			StartRecord=StartRecord+1
		end if

		'how many rows per page? 
		RowsPerPage = aspL.convertNmbr(aspL.getRequest("length"))
		if RowsPerPage = 0 then RowsPerPage=10
		 
		'reading search phrase - this one may be empty
		strSearch = trim(aspL.getRequest("search[value]"))	
		
		aspL.aspError("cls_dt_returnJson.Class_Initialize")
		
		on error goto 0
	
	end sub
	
	public sub dumpJson
	
		on error resume next
		
		'the recordset rs expects an open database connection (db)
		'this is a very basic SQL implementation. 
		'If you need Group By/having, you would need to add it here
		set rs=db.rs : rs.Open strSelect & strWhere  & strOrder
		
		'total number of results
		rTotal=rs.recordcount
		
		'we only want a portion of the recordset 
		'starting from the startrecord (AbsolutePosition), and the next x rows (pagesize) only
		'if there are no records returned at all, do not set AbsolutePosition as this raises an error
		'no need for pagesize either in that case
		if rTotal>0 then
			rs.AbsolutePosition=StartRecord
			rs.pagesize=RowsPerPage
		end if

		'prepare JSON return - this class takes care of the recordset paging! - see json.recordsetPaging
		json.recordsetPaging=true
		JsonAnswer=json("data", rs, false) 

		'finalizing JSON response - preparing header:
		JsonHeader = "{ ""draw"": "& draw &", "& vbcrlf
		JsonHeader = JsonHeader & """recordsTotal"": " & rTotal &", "& vbcrlf
		JsonHeader = JsonHeader & """recordsFiltered"": " & rTotal &", "& vbcrlf
		  
		'removing from generated JSON initial bracket { and concatenating all together.
		JsonAnswer=right(JsonAnswer,Len(JsonAnswer)-1)
		JsonAnswer = JsonHeader & JsonAnswer

		set db=nothing : set rs=nothing
		 
		aspL.aspError("cls_dt_returnJson.dumpJson")
		
		'writing a response and stop executing page
		aspL.dumpJson JsonAnswer
		
		on error goto 0
	
	end sub
	
end class
%>