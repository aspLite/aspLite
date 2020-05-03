<%
'CONDITIONAL load of asp page = include file
aspL.exec("code/demo_asp/class.asp")	

dim testObj
set testObj=new cls_test

'formbuilder sample - built on bootstrap css
aspl.exec("code/demo_asp/formbuilder/formbuilder.asp")

dim form : set form=new cls_formbuilder
form.listenTo "e","sampleform10"
form.targetDiv="sampleform10"
form.id="conditionalinclude"
	
'result
dim feedback : set feedback=form.field
feedback.add "type","comment"
feedback.add "html",testObj.hello
feedback.add "tag","div"

form.build()

%>