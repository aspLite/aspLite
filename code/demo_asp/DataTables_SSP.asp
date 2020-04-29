<%
'load the cls_dt_returnJson class that will return the Json at the end 
'an object returnJson is already created in that asp file
aspL.exec("code/demo_asp/datatables/returnJson.asp")

'pass the selectClause to the returnJson object
returnJson.strSelect="select * from testbigdata "

'if a searchstring is entered, you have to decide on what to search for exactly
'make sure to ALWAYS use aspl.sqli!!! (or anything else to prevent SQLinjection)
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

'finally return the JSon object to DataTables - here stops the page execution
returnJson.dumpJson()

%>