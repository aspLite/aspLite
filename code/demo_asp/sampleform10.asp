<%
'CONDITIONAL load of asp page = include file
aspL.exec("code/demo_asp/class.asp")	

dim testObj
set testObj=new cls_test

dim form : set form=new cls_formbuilder
form.targetDiv="sampleform10"
form.id="conditionalinclude"
	
'result
dim feedback : set feedback=form.field
feedback.add "type","element"
feedback.add "html",testObj.hello
feedback.add "tag","div"

form.build()

%>