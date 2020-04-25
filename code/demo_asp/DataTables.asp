<%
body=aspL.loadText("html/demo_asp/DataTables.resx")

dim i, randomizer, table
set randomizer=aspL.plugin("randomizer")

table="<thead>"
table=table & "<tr>"
table=table & "<th>col 1</th>"
table=table & "<th>col 2</th>"
table=table & "<th>col 3</th>"
table=table & "<th>col 4</th>"
table=table & "</tr>"
table=table & "</thead>"

table=table & "<tbody>"

for i=1 to 250
	'generate some random content for the table cells
	table=table & "<tr>"
		table=table & "<td>" & randomizer.randomnumber(10000,99999) & "</td>"
		table=table & "<td>" & randomizer.randomtext(randomizer.randomnumber(7,10)) & "</td>"
		table=table & "<td>" & randomizer.randomtext(randomizer.randomnumber(7,10)) & "</td>"	
		table=table & "<td>" & randomizer.randomtext(randomizer.randomnumber(7,10)) & "</td>"
	table=table & "</tr>"
next

table=table & "</tbody>"

set randomizer=nothing

body=replace(body,"[datatable]",table,1,-1,1)	
%>