<%
'hello world plugin example
dim helloworld
set helloworld=aspL.plugin("helloworld")

dim form : set form=aspl.form
form.target="sampleform9"
	
'result
dim feedback : set feedback=form.field("element")
feedback.add "html",helloworld.hw()
feedback.add "tag","div"

form.build()
%>