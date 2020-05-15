<%
'aspLite randomizer example
dim i

for i=1 to 10
	'generate some random words with random lengths (10-20)
	body=body & aspL.randomizer.randomtext(aspL.randomizer.randomnumber(5,10)) & " "
next

dim form : set form=aspl.form
	
'result
dim feedback : set feedback=form.field("element")
feedback.add "html",body
feedback.add "tag","div"

form.build()

%>