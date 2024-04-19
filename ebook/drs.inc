<%
dim form : set form=aspl.form

dim rs : set rs=aspl.recordset

rs.fields.append "firstname",203,70
rs.fields.append "lastname",203,60
rs.fields.append "birthdate",133,25
rs.fields.append "description",203,25
rs.fields.append "admin",11

rs.open

dim i
for i=1 to 10
	rs.AddNew
	rs("firstname")		= aspl.pcase(aspl.randomizer.randomtext(5))
	rs("lastname")		= aspl.pcase(aspl.randomizer.randomtext(7))
	rs("birthdate")		= dateadd("d",(round(aspl.randomizer.randomNumber(10000,25000),0))*-1,date())
	rs("description")	= aspl.pcase(aspl.randomizer.randomtext(10))
	rs("admin")	= true
	rs.Update
next

'always movefirst before exporting the recordset! 
'by default, after a loop like the one we do, the recordset is EOF (end of file)!
rs.movefirst 

'apply sorting
rs.sort="lastname asc, firstname asc"

'recordcount returns the total number of records in a recordset
form.write "We just added <strong>" & rs.recordcount & "</strong> records."

form.newline
form.write "aspl.json.toJSON converts this recordset to JSON:"
form.newline

form.writejs "$('#drs').append('"& aspl.json.toJSON("recordset",rs,true) & "');"

form.build
%>