<%
'formbuilder sample - built on bootstrap css
aspl.exec("code/demo_asp/formbuilder/formbuilder.asp")

dim form : set form=new cls_formbuilder
form.listenTo "e","rate"
form.targetDiv="body"
form.requiredLegend=""
form.requiredStar=""
form.id="rateForm"

'form-submitted
if form.postback then
	
	'feedback
	dim feedback : set feedback=aspl.dict
	feedback.add "html","Thanks so much for your " & aspL.sanitize(aspL.getRequest("score")) & " stars!"
	feedback.add "type","comment"
	feedback.add "tag","div"
	form.addField(feedback)	
	
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

dim score : set score=aspl.dict
score.add "label","How would you rate aspLite so far?"
score.add "type","select"
score.add "name","score"
score.add "class","form-control"
score.add "id","score"
score.add "options",stars
score.add "onchange","$('#" & form.id & "').submit();"

form.addField(score)

'##########################################################################################

form.build()

%>