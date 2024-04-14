<%
dim form : set form=aspl.form

dim db : set db=aspL.plugin("database")

db.dbms=1 'Access - default
'db.dbms=2: SQL Server
'db.dbms=3: MySQL

db.path="ebook/access.mdb" 

'Read/Write Recordset
dim rs : set rs=db.rs
rs.open "select * from person where 1=2"
rs.addNew()
rs("lastname")=aspl.pCase(aspl.randomizer.randomText(10))
rs("firstname")=aspl.pCase(aspl.randomizer.randomText(10))
rs.update
rs.close
set rs=nothing

'Read-only recordset
set rs=db.execute("select top 10 * from person")

while not rs.eof
	form.write rs("lastname") & ",&nbsp;" & rs("firstname") & "<br>"
	rs.movenext
wend

form.newline
form.write "If you want to know the total number of records: "
form.write "&nbsp;<strong>" & db.execute("select count(*) from person")(0) & "</strong>"

form.newline
form.write "If you want to use recordcount however, you need a read/write recordset again: "

set rs=db.rs : rs.open("select * from person")
form.write "&nbsp;<strong>" & rs.recordcount & "</strong>"

set rs=nothing
set db=nothing

form.build

%>