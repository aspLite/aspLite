<%
dim form : set form=aspl.form
form.listenTo "asplEvent","sampleform17"

'rather than set the form-target, you can also make it dynamic,
'so that the same form can bind to multiple div-ids
'sampleForm17 binds to both id=sampleform17 and id=sampleform19
form.target=aspl.getRequest("asplTarget")

'in that case, make sure to add the asplTarget-field to the form
set asplTarget=form.field("hidden")
asplTarget.add "name","asplTarget"
asplTarget.add "value",aspl.getRequest("asplTarget")

if form.postback then

	dim feedback : set feedback=form.field("element")
	feedback.add "tag","div"
	feedback.add "html","Thank you for subscribing!"
	feedback.add "class","alert alert-warning"	
	
	form.build()

end if 
	
'result
dim email : set email=form.field("email")
email.add "name","email"
email.add "class","form-control"
email.add "placeholder","your email"
email.add "required",true

dim submit : set submit=form.field("submit")
submit.add "value","Subscribe"
submit.add "class","btn btn-info"
submit.add "style","margin-top:5px"

form.build()

%>