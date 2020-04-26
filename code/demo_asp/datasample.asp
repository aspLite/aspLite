<%
on error resume next

'database plugin example
dim db : set db=aspL.plugin("database") : db.path="db/sample.mdb"
dim rs : set rs=db.execute("select * from person")

body=body & "Access database-plugin:<ul>"

while not rs.eof
	body=body & "<li>" & rs("sName") & "</li>"
	rs.movenext
wend 

body=body & "</ul>"

on error goto 0
%>