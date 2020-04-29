<%
class cls_dt_returnJson

	private OrderCol,OrderDir,strOrder,draw,StartRecord,RowsPerPage,JsonAnswer,JsonHeader
	public strSearch, strWhere, dbPath, strSelect, db

	Private Sub Class_Initialize()
	
		'which column will have to be sorted on ?
		OrderCol = aspL.convertNmbr(aspL.getRequest("Order[0][column]"))
		
		'asc or desc?
		OrderDir =  trim(lcase(aspL.getRequest("Order[0][dir]")))
		if OrderDir<>"asc" and OrderDir<>"desc" then OrderDir="asc"
		 
		'WHERE clause uses columns number, like e.g: ORDER BY 1 DESC, you may add translations to column names here, like e.g.: OrderCol = Replace(OrderCol,"0","Col1")
		'We are adding 1 here, because DataTables indexes columns starting from 0
		strOrder=" Order By " & OrderCol+1 & " " & OrderDir
		 
		draw = aspL.convertNmbr(aspL.getRequest("draw"))
		
		'where exactly are in the table?
		StartRecord = aspL.convertNmbr(aspL.getRequest("start"))

		if StartRecord = 0 then 
			StartRecord=1
		else
			StartRecord=StartRecord+1
		end if

		'how many rows per page? 
		RowsPerPage = aspL.convertNmbr(aspL.getRequest("length"))
		if RowsPerPage = 0 then RowsPerPage=10
		 
		'reading search phrase - this one may be empty
		strSearch = trim(aspL.getRequest("search[value]"))	
	
	end sub
	
	public sub dumpJson
		
		'the recordset rs expects an open database connection (db)
		set rs=db.rs : rs.Open strSelect & strWhere  & strOrder
		
		'total number of results
		rTotal=rs.recordcount
		
		'we only want a portion of the recordset 
		'starting from the startrecord (AbsolutePosition), and next x rows (pagesize)
		if rTotal>0 then
			rs.AbsolutePosition=StartRecord
			rs.pagesize=RowsPerPage
		end if

		'prepare JSON return - this class takes care of the paging! - see jsonObj.recordsetPaging
		dim jsonObj : set jsonObj=aspL.plugin("json") : jsonObj.recordsetPaging=true
		JsonAnswer=jsonObj.toJSON("data", rs, false) 

		'finalizing JSON response - preparing header:
		JsonHeader = "{ ""draw"": "& draw &", "& vbcrlf
		JsonHeader = JsonHeader & """recordsTotal"": " & rTotal &", "& vbcrlf
		JsonHeader = JsonHeader & """recordsFiltered"": " & rTotal &", "& vbcrlf
		  
		'removing from generated JSON initial bracket { and concatenating all together.
		JsonAnswer=right(JsonAnswer,Len(JsonAnswer)-1)
		JsonAnswer = JsonHeader & JsonAnswer

		set db=nothing : set rs=nothing : set jsonObj=nothing
		 
		'writing a response and stop executing page
		aspL.dumpJson(JsonAnswer)	
	
	end sub
	
end class
%>