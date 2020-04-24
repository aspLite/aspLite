<%
on error resume next

Response.LCID = 1033

'load the plugincode in the namespace of this page
aspL.plugin("json")

dim db : set db=aspl.plugin("database")
db.path="db/sample.mdb"
dim rs : set rs=db.getDynmicRS 
rs.open ("select * from person")

set jsonArr=new JSONarray
jsonArr.LoadRecordSet rs

aspL.asperror("json2html")

aspL.dump "{""records"":" & jsonArr.Serialize & "}"

on error goto 0
%>