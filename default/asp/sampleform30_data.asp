<%
on error resume next

dim rs : set rs=db.rs : rs.open("select top 1500 * from contact order by iId asc")

aspL.asperror("json_datatables")

aspl.json.dump(rs)

on error goto 0
%>