<!-- #include file="asplite/asplite.asp"-->
<%

dim db,rs,field,datatable
	
set db=aspL.plugin("database") : db.path="db/sample.mdb" : set rs=db.rs

rs.open ("delete from contact")

rs.open ("select * from contact where 1=2")

dim i
for i=0 to 9999
	rs.AddNew
	rs("sText") = aspl.pCase(aspl.randomizer.randomText(6)) & " " & aspl.pCase(aspl.randomizer.randomText(11))
	rs("iNumber") = aspl.randomizer.randomNumber(0,9999)
	rs("dDate") = DateSerial(aspl.randomizer.randomNumber(1970,2020),aspl.randomizer.randomNumber(1,12),aspl.randomizer.randomNumber(1,28))
	rs("iCountryID") = aspl.randomizer.randomNumber(1,248)
	rs.Update

next

rs.close
set rs=nothing

%>