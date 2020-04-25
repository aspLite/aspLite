<%
on error resume next

Response.LCID = 1033
Response.ContentType = "application/json"

set db=aspL.plugin("database")
db.path="db/sample.mdb"
set rs=db.execute("select * from testdata")

'load the JSON plugin in the namespace of this page
aspL.plugin("json")
set jsonArr=new JSONarray
jsonArr.LoadRecordSet rs

aspL.asperror("json_datatables")

aspL.dump "{""aspLrecords"":" & jsonArr.Serialize & "}"

on error goto 0
%>