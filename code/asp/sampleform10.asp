<%
'CONDITIONAL load of asp page = include file
aspL.exec("code/asp/class.asp")	

dim testObj
set testObj=new cls_test

dim form : set form=aspl.form
form.target="sampleform10"
	
'result
dim feedback : set feedback=form.field("element")
feedback.add "html",testObj.hello
feedback.add "tag","div"

form.build()

%>