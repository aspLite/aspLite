<%
'hello world plugin example
dim helloworld
set helloworld=aspL.plugin("helloworld")

'formbuilder sample - built on bootstrap css
aspl.exec("code/demo_asp/formbuilder/formbuilder.asp")

dim form : set form=new cls_formbuilder
form.listenTo "e","sampleform9"
form.targetDiv="sampleform9"
form.id="helloworld"
	
'result
dim feedback : set feedback=form.field
feedback.add "type","comment"
feedback.add "html",helloworld.hw()
feedback.add "tag","div"

form.build()

%>