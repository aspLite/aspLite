<%
'hello world plugin example
dim helloworld
set helloworld=aspL.plugin("helloworld")

dim form : set form=aspl.form
	
'result
dim feedback : set feedback=form.field("element")
feedback.add "html",helloworld.hw()
feedback.add "tag","div"

form.build()
%>