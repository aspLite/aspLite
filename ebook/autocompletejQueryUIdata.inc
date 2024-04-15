<%
'we expect a parameter named "term" from the ajax call. make sure to protect against SQLinjection

dim output, db, rs
set db=aspl.plugin("database") : db.path="ebook/access.mdb"

set rs=db.rs
rs.open("select top 100, lastname from person where lastname like '%" & aspl.sqli(aspl.getRequest("term")) & "%'")

output = "["

while not rs.eof

	output = output & """" & aspl.json.escape(rs(0)) & ""","
	rs.movenext
	
wend 

output = left(output,len(output)-1) & "]"

aspl.dumpJson(output)
%>