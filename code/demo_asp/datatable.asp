<%
dim db,rs,field,datatable
	
set db=aspL.plugin("database")
db.path="db/sample.mdb"
set rs=db.GetDynamicRS

rs.open ("select * from person")

datatable="<tr>"

'get headers
for each field in rs.fields	
	datatable=datatable & "<th scope=""col"">" & aspL.sanitize(field.name) & "</th>"		
next

datatable=datatable & "</tr>"

while not rs.eof	

	datatable=datatable & "<tr>"

	for each field in rs.fields	
		datatable=datatable & "<td>" & aspL.sanitize(rs(field.name)) & "</td>"
	next
	
	datatable=datatable & "</tr>"
	
	rs.movenext

wend

set db=nothing
set rs=nothing

aspL.dump "<table class=""table table-striped"">" & datatable & "</tr>"

%>