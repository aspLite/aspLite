<%
on error resume next

'as we're loading 1000 records, we use the aspLite caching feature. It only refreshes the "testdata" every 10 minutes (600 seconds)

dim cachedJson
cachedJson=aspL.getCacheT("testdata",600)

if aspL.isEmp(cachedJson) then

	dim db : set db=aspL.plugin("database")
	db.path="db/sample.mdb"
	dim rs : set rs=db.execute("select * from testdata")

	aspL.plugin("json")
	dim jsonArr : set jsonArr=new JSONarray
	jsonArr.LoadRecordSet rs
	
	cachedJson=jsonArr.Serialize
	
	set jsonArr=nothing
	set db=nothing
	set rs=nothing
	
	aspL.setCache "testdata",cachedJson

end if

aspL.asperror("json_datatables")

aspL.dumpJson "{""aspLrecords"":" & cachedJson & "}"

on error goto 0
%>