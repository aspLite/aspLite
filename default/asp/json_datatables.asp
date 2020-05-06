<%
on error resume next

dim rs : set rs=db.rs : rs.open("select * from testdata")

aspL.asperror("json_datatables")

aspl.json.dump(rs)

on error goto 0
%>