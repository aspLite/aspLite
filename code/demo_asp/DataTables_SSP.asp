<%
'solution as found on https://datatables.net/forums/discussion/59746/new-classic-asp-server-side-script-for-datatables-1-10-20

on error resume next
aspL.exec("code/demo_asp/json/json.asp")

'Reading ordering data nd settin to empty if none. Empty should never happen, but it's just for safety.
OrderCol = aspL.convertNmbr(Request("Order[0][column]"))
OrderDir =  aspL.getRequest("Order[0][dir]")
if not OrderCol = 0 and not OrderDir = "" then
 
	'WHERE clause uses columns number, like e.g: ORDER BY 1 DESC, you may add translations to column names here, like e.g.: OrderCol = Replace(OrderCol,"0","Col1")
	'We are adding 1 here, because DataTables indexes columns starting from 0
	strOrder=" Order By " & OrderCol+1 & " " & OrderDir

end if
 
'reading numbers sent by DataTables and setting them to defaults in case of empty (which should never happen):
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
strSearch = (aspL.getRequest("search[value]"))
 
'if not empty, then gerenate 'WHERE' Clause. Here you should edjust query to your DB.
if not aspL.isEmp(strSearch) then

	strWhere = " where text like '%" & aspl.sqli(strSearch) & "%'"	
	
	if aspL.convertNmbr(strSearch)<>0 then
		'this would only work in Access
		strWhere=strWhere & " or id="& strSearch
		strWhere=strWhere & " or number="& strSearch
		strWhere=strWhere & " or year(date)="& strSearch
		strWhere=strWhere & " or month(date)="& strSearch
		strWhere=strWhere & " or day(date)="& strSearch		
	end if
	
else
	strWhere=""
end if

dim db : set db=aspL.plugin("database")
db.path="db/sample.mdb"
set rsReport=db.getDynamicRS

'first query to get data from DB to table with WHERE, ORDER and 'PAGINATION' clauses
strReport = "select * from testbigdata " & strWhere  & strOrder 

rsReport.Open strReport
rTotal=rsReport.recordcount

if rTotal>0 then
	rsReport.AbsolutePosition=StartRecord
	rsReport.pagesize=RowsPerPage
end if

dim jsonObj
set jsonObj=new json
jsonObj.recordsetPaging=true

JsonAnswer=jsonObj.toJSON("data", rsReport, false) 
 
'finalizing JSON response - preparing header:
JsonHeader = "{ ""draw"": "& draw &", "& vbcrlf
JsonHeader = JsonHeader & """recordsTotal"": " & rTotal &", "& vbcrlf
JsonHeader = JsonHeader & """recordsFiltered"": " & rTotal &", "& vbcrlf
  
'removing from generated JSON initial bracket { and concatenating all toghether.
JsonAnswer=right(JsonAnswer,Len(JsonAnswer)-1)
JsonAnswer = JsonHeader & JsonAnswer
 
'writing a response:
aspL.dumpJson(JsonAnswer)
%>