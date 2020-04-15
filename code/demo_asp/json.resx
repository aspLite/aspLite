<%
on error resume next

Response.LCID = 1033

'simply load the plugincode in the namespace of this asp script
aspL.plugin("json")

dim db
set db=aspl.plugin("accessDB")
db.path="db/sample.mdb"

set rs=db.GetDynamicRS

rs.open ("select * from person")

set jsonArr=new JSONarray
jsonArr.LoadRecordSet rs

dim output
output=jsonArr.Serialize	

set jsonArr=nothing
set rs=nothing
set db=nothing

aspL.asperror("json")

aspL.dump "{""records"":" & output & "}"

on error goto 0
%>