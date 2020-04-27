<%
on error resume next

dim db : set db=aspL.plugin("database") : db.path="db/sample.mdb"
dim rs : set rs=db.rs : rs.open ("select * from person")

dim jsonObj : set jsonObj=aspL.plugin("json")
dim JsonAnswer : JsonAnswer=jsonObj.toJSON("aspLrecords", rs, false)	

set jsonObj=nothing : set db=nothing : set rs=nothing

aspL.asperror("json")

aspL.dumpJson JsonAnswer

on error goto 0
%>