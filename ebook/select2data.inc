<%
'we expect a parameter named "term " from the ajax call. make sure to protect against SQLinjection
'in this query, the searchstring will be searched for in the entire name ('%term %'). 
'if you prefer to only search for person names BEGINNING with the querystring, use 'term %'

dim db : set db=aspl.plugin("database")
db.path="ebook/access.mdb"

dim suggestions

if len(trim(aspl.getRequest("term")))>0 then

	dim cleanTerm : cleanTerm=trim(aspl.removeTabs(aspl.sqli(aspl.getRequest("term"))))
	
	sql="select * from person  "
	sql=sql & " where firstname like '%" & cleanTerm & "%' or lastname like '%" & cleanTerm & "%' "	
	sql=sql & " order by lastname, firstname"
	
	set rs=db.rsOpen(sql)
	
	while not rs.eof
		suggestions=suggestions & "{ ""text"": """ & aspl.json.escape(rs("lastname")) & ", "
		suggestions=suggestions & aspl.json.escape(rs("firstname")) & """, ""id"": """ & rs("id") & """ },"
		rs.movenext
	wend 
	
	set rs=nothing
	
	suggestions=suggestions & "{ ""text"": ""No results"", ""id"": ""0"" },"
	
end if

if right(suggestions,1)="," then
	suggestions=left(suggestions,len(suggestions)-1)
end if

aspl.dumpJson("{""results"": [" & suggestions & "]}")

%>