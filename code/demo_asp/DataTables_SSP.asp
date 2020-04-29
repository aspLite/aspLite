<%
on error resume next

'load the cls_dt_returnJson class that will return the Json in the end 
aspL.exec("code/demo_asp/datatables/returnJson.asp")

dim returnJson : set returnJson=new cls_dt_returnJson

'pass the selectClause to returnJson
returnJson.strSelect="select * from testbigdata "

'if a searchstring is entered, you have to decide on what to search for exactly
'make sure to use aspl.sqli!!! (prevents SQLinjection)
if not aspL.isEmpty(returnJson.strSearch) then

	strWhere = " where sText like '%" & aspl.sqli(returnJson.strSearch) & "%' "	
	
	'in this case I would like to search on numeric/date fields as well
	'but first check if strSearch is numeric with aspL.isNumber
	
	if aspL.isNumber(returnJson.strSearch) then
		'this would only work in Access (the date functions)
		strWhere=strWhere & " or iId="& returnJson.strSearch
		strWhere=strWhere & " or iNumber="& returnJson.strSearch
		strWhere=strWhere & " or year(dDate)="& returnJson.strSearch
		strWhere=strWhere & " or month(dDate)="& returnJson.strSearch
		strWhere=strWhere & " or day(dDate)="& returnJson.strSearch		
	end if
	
	'return the whereClause to returnJson	
	returnJson.strWhere=strWhere
	
end if

'connect to the database- in this case an access database (dbpath)
set returnJson.db=aspL.plugin("database")
returnJson.db.path="db/sample.mdb"

'finally return the JSon object to DataTables
returnJson.dumpJson()

%>