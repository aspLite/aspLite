<%
'randomizer plugin example
dim i

for i=1 to 10
	'generate some random words with random lengths (10-20)
	body=body & aspL.plugin("randomizer").randomtext(aspL.plugin("randomizer").randomnumber(5,10)) & " "
next

dim form : set form=new cls_formbuilder
form.targetDiv="sampleform11"
form.id="randomizer"
	
'result
dim feedback : set feedback=form.field
feedback.add "type","comment"
feedback.add "html",body
feedback.add "tag","div"

form.build()

%>