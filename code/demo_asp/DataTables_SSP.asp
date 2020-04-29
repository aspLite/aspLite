<%
'solution as found on https://datatables.net/forums/discussion/59746/new-classic-asp-server-side-script-for-datatables-1-10-20

on error resume next

OrderCol = aspL.convertNmbr(aspL.getRequest("Order[0][column]"))
OrderDir =  lcase(aspL.getRequest("Order[0][dir]"))

if OrderDir<>"asc" and OrderDir<>"desc" then OrderDir="asc"
 
'WHERE clause uses columns number, like e.g: ORDER BY 1 DESC, you may add translations to column names here, like e.g.: OrderCol = Replace(OrderCol,"0","Col1")
'We are adding 1 here, because DataTables indexes columns starting from 0
strOrder=" Order By " & OrderCol+1 & " " & OrderDir
 
draw = aspL.convertNmbr(aspL.getRequest("draw"))
 
StartRecord = aspL.convertNmbr(aspL.getRequest("start"))

if StartRecord = 0 then 
	StartRecord=1
else
	StartRecord=StartRecord+1
end if

RowsPerPage = aspL.convertNmbr(aspL.getRequest("length"))
if RowsPerPage = 0 then RowsPerPage=10
 
'reading search phrase - this one may be empty
strSearch = aspL.getRequest("search[value]")
 
'if not empty, then gerenate 'WHERE' Clause. Here you should edjust query to your DB.
if not aspL.isEmpty(strSearch) then

	strWhere = " where sText like '%" & aspl.sqli(strSearch) & "%'"	
	
	if aspL.convertNmbr(strSearch)<>0 then
		'this would only work in Access
		strWhere=strWhere & " or iId="& strSearch
		strWhere=strWhere & " or iNumber="& strSearch
		strWhere=strWhere & " or year(dDate)="& strSearch
		strWhere=strWhere & " or month(dDate)="& strSearch
		strWhere=strWhere & " or day(dDate)="& strSearch		
	end if
	
end if

dim db : set db=aspL.plugin("database") : db.path="db/sample.mdb"
set rs=db.rs : rs.Open "select * from testbigdata " & strWhere  & strOrder

rTotal=rs.recordcount

if rTotal>0 then
	rs.AbsolutePosition=StartRecord
	rs.pagesize=RowsPerPage
end if

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
 
'writing a response:
aspL.dumpJson(JsonAnswer)
%>