<%
on error resume next

dim rs : set rs=db.rs : rs.open ("select * from person")

aspL.asperror("json")

aspl.json.dump(rs)	

on error goto 0
%>