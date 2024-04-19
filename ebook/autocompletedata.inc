<%
'we expect a parameter named "query" from the ajax call. make sure to protect against SQLinjection
'in this query, the searchstring will be searched for in the entire country-name ('%query%'). 
'if you prefer to only search for persons lastnames BEGINNING with the querystring, use 'query%'

dim db : set db=aspl.plugin("database")
db.path="ebook/access.mdb"

dim rs : set rs=db.rs
dim sql : sql="select top 100, lastname from person "
sql=sql & " where lastname like '%" & aspl.sqli(aspl.getRequest("query")) & "%' order by lastname asc"

rs.open(sql)

if not rs.eof then
	aspl.dumpJson(aspl.json.toJSON("suggestions",rs.getRows,false))
else
	aspl.dumpJson(aspl.json.toJSON("suggestions","",false))
end if
%>