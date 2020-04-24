<%
on error resume next

Response.LCID = 1033
Response.ContentType = "application/json"

'load the plugincode in the namespace of this page
aspL.plugin("json")

dim db : set db=aspl.plugin("database")
db.path="db/sample.mdb"
dim rs : set rs=db.execute ("select * from person")

set jsonArr=new JSONarray
jsonArr.LoadRecordSet rs

aspL.asperror("json")

aspL.dump "{""data"":" & jsonArr.Serialize & "}"

on error goto 0
%>