<%
on error resume next

dim db : set db=aspL.plugin("database") : db.path="db/sample.mdb"
dim rs : set rs=db.rs : rs.open("select * from testdata")

aspL.asperror("json_datatables")

json.dump(rs)

on error goto 0
%>