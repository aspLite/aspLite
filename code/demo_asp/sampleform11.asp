<%
'randomizer plugin example
dim i

for i=1 to 10
	'generate some random words with random lengths (10-20)
	body=body & aspL.plugin("randomizer").randomtext(aspL.plugin("randomizer").randomnumber(5,10)) & " "
next

'formbuilder sample - built on bootstrap css
aspl.exec("code/demo_asp/formbuilder/formbuilder.asp")

dim form : set form=new cls_formbuilder
form.listenTo "e","sampleform10"
form.targetDiv="sampleform11"
form.id="randomizer"
	
'result
dim feedback : set feedback=form.field
feedback.add "type","comment"
feedback.add "html",body
feedback.add "tag","div"

form.build()

%>