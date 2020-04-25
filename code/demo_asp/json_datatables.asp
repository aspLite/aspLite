<%
on error resume next

Response.LCID = 1033
Response.ContentType = "application/json"

dim rs
set rs = createObject("ADODB.Recordset")

rs.Fields.Append "id", 3, , adFldKeyColumn
rs.Fields.Append "Firstname", 200, 50, adFldMayBeNull
rs.Fields.Append "Lastname", 200, 50, adFldMayBeNull

rs.Open

dim i
for i=1 to 250
	rs.AddNew
	rs("id") = i
	rs("Firstname") = aspl.plugin("randomizer").randomText(10)
	rs("Lastname") = aspl.plugin("randomizer").randomText(15)
	rs.Update
next

rs.moveFirst

'load the JSON plugin in the namespace of this page
aspL.plugin("json")
set jsonArr=new JSONarray
jsonArr.LoadRecordSet rs

rs.Close

set rs = nothing

aspL.asperror("json_datatables")

aspL.dump "{""aspLrecords"":" & jsonArr.Serialize & "}"

on error goto 0
%>