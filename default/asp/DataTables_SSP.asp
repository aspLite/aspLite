<%
on error resume next

'load the cls_dt_returnJson class that will return the Json at the end 
'an object returnJson is already created in that asp file
aspL("default/asp/datatables/returnJson.asp")


'pass the selectClause to the returnJson object
'to make this a little more interesting, i added a linked table (left join)
'GOOD PRACTICE: avoid select * as this will mess up the order of the ordered fields!
'ordered fields are represented by a number, in this example starting with 0 (contact.iId) to 5 (countryname)
'Also, note that both tables contact and country have a field named "sText". To avoid problems, we alias to countryName
strSelect="SELECT contact.iId, contact.sText, contact.iNumber, contact.bBoolean, contact.dDate, country.sText as countryName "
strSelect=strSelect & "FROM contact LEFT JOIN country ON contact.iCountryID = country.iId"

'pass the selectClause to the returnJson object
returnJson.strSelect=strSelect

'if a searchstring is entered, you have to decide on what to search for exactly
'make sure to ALWAYS use aspl.sqli!!! (or anything else to prevent SQLinjection)
if not aspL.isEmpty(returnJson.strSearch) then

	'search in the contact-table
	strWhere = " where contact.sText like '%" & aspl.sqli(returnJson.strSearch) & "%' "	
	
	'but also in the country table
	strWhere = strWhere & " or country.sText like '%" & aspl.sqli(returnJson.strSearch) & "%' "
		
	'in this case I would like to search on numeric/date fields as well
	'but first check if strSearch is numeric with aspL.isNumber	
	if aspL.isNumber(returnJson.strSearch) then			
		strWhere=strWhere & " or contact.iNumber="& returnJson.strSearch
		'this would only work in Access (the date functions)	
		strWhere=strWhere & " or year(contact.dDate)="& returnJson.strSearch
		strWhere=strWhere & " or month(contact.dDate)="& returnJson.strSearch
		strWhere=strWhere & " or day(contact.dDate)="& returnJson.strSearch		
	end if
	
	'return the whereClause to returnJson	
	returnJson.strWhere=strWhere
	
end if

aspL.aspError("datatables server-side-processing")

'finally return the JSon object to DataTables - here stops the page execution
returnJson.dumpJson()

%>