<%

dim form : set form=aspl.form
form.listenTo "e","sampleform17"
form.target="sampleform17"

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