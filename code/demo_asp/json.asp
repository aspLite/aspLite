<%
on error resume next

dim db : set db=aspL.plugin("database") : db.path="db/sample.mdb"
dim rs : set rs=db.rs : rs.open ("select * from person")

aspL.asperror("json")

json.dump(rs)	

on error goto 0
%>