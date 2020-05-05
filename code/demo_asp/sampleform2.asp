<%

dim form : set form=new cls_formbuilder
form.listenTo "e","sampleform2"
form.targetDiv="sampleform2"
form.id="rateForm"

'form-submitted
if form.postback then

	form.doScroll=true
	
	'feedback
	dim feedback : set feedback=form.field
	feedback.add "html","Thanks so much for your <strong>" & aspL.convertNmbr(aspL.getRequest("score")) & " stars</strong>!"
	feedback.add "type","element"
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

dim score : set score=form.field
score.add "label","How would you rate aspLite so far?"
score.add "type","select"
score.add "name","score"
score.add "id","score"
score.add "class","form-control"
score.add "options",stars 'VBSscript dictionary!
score.add "onchange","$('#" & form.id & "').submit();"

'##########################################################################################

form.build()

%>