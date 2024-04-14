<%
dim form : set form=aspl.form
dim db : set db=aspl.plugin("database")
db.path="ebook/access.mdb"

dim i,table

table="<table class=""table table-striped"" id=""myTable"">"
table=table & "<thead>"
table=table & "<tr>"
table=table & "<th>ID</th>"
table=table & "<th>First name</th>"
table=table & "<th>Last name</th>"
table=table & "</tr>"
table=table & "</thead>"
table=table & "<tbody>"

dim rs : set rs=db.execute("select top 250, * from person order by id desc")

while not rs.eof

	table=table & "<tr>"
	table=table & "<td>" & rs("id") & "</td>"
	table=table & "<td>" & rs("firstname") & "</td>"
	table=table & "<td>" & rs("lastname") & "</td>"
	table=table & "</tr>"

	rs.movenext
wend

table=table & "</tbody>"
table=table & "</table>"

form.write table

form.writejs "$(document).ready(function() {$('#myTable').DataTable()});"

form.build

%>