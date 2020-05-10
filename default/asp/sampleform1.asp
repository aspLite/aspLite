<%

dim form : set form=aspl.form

'form-submitted
if form.postback then

	form.doScroll=true
	
	'feedback
	dim feedback : set feedback=form.field("element")
	feedback.add "html","Hello " & aspL.sanitize(aspL.getRequest("yourname")) & "!"
	feedback.add "tag","div"
	feedback.add "class","alert alert-warning"	
	
	'you can build the form here already. This will stop further exection	
	form.build()

end if	

'##########################################################################################

dim yourname : set yourname=form.field("text")
yourname.add "placeholder","Your name:"
yourname.add "name","yourname"
yourname.add "class","form-control"
yourname.add "maxlength",50
yourname.add "required",true

'##########################################################################################

dim submit : set submit=form.field("submit")
submit.add "style","margin-top:10px"
submit.add "name","btnAction"
submit.add "value","Submit"
submit.add "class","btn btn-primary"

'##########################################################################################

form.build()

%>