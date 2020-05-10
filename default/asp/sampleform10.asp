<%
'CONDITIONAL load of asp page = include file
aspL("default/asp/includes/class.asp")	

dim testObj
set testObj=new cls_test

dim form : set form=aspl.form
	
'result
dim feedback : set feedback=form.field("element")
feedback.add "html",testObj.hello
feedback.add "tag","div"
feedback.add "class","jumbotron"

form.build()

%>