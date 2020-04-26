<%
on error resume next

dim db : set db=aspL.plugin("database")
db.path="db/sample.mdb"
dim rs : set rs=db.execute ("select * from person")

'load the json-plugin in the namespace of this page
aspL.plugin("json")
set jsonArr=new JSONarray
jsonArr.LoadRecordSet rs

aspL.asperror("json")

aspL.dumpJson "{""aspLrecords"":" & jsonArr.Serialize & "}"

on error goto 0
%>