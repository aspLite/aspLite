<%

dim form : set form=aspl.form
form.listenTo "e","sampleform2"
form.target="sampleform2"

'form-submitted
if form.postback then

	form.doScroll=true
	
	'feedback
	dim feedback : set feedback=form.field("element")
	feedback.add "html","Thanks so much for your <strong>" & aspL.convertNmbr(aspL.getRequest("score")) & " stars</strong>!"
	feedback.add "tag","div"
	
	'you can build it here already. This will stop further exection	
	form.build()

end if	

'##########################################################################################

dim stars : set stars=aspl.dict
stars.add 1, "1 star"
stars.add 2, "2 stars"
stars.add 3, "3 stars"
stars.add 4, "4 stars"
stars.add 5, "5 stars"

dim score : set score=form.field("select")
score.add "label","How would you rate aspLite so far?"
score.add "name","score"
score.add "id","score"
score.add "class","form-control"
score.add "emptyfirst","please select"
score.add "options",stars 'VBSscript dictionary!
score.add "onchange","$('#" & form.id & "').submit();"

'##########################################################################################

form.build()

%>