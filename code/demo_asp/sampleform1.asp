<%

dim form : set form=new cls_formbuilder
form.listenTo "e","sampleform1"
form.targetDiv="sampleform1"
form.id="helloForm"

'form-submitted
if form.postback then

	form.doScroll=true
	
	'feedback
	dim feedback : set feedback=form.field
	feedback.add "type","element"
	feedback.add "html","Hello " & aspL.sanitize(aspL.getRequest("yourname")) & "!"
	feedback.add "tag","div"
	feedback.add "class","alert alert-warning"	
	
	'you can build it here already. This will stop further exection	
	form.build()

end if	

'##########################################################################################

dim yourname : set yourname=form.field
yourname.add "placeholder","Your name:"
yourname.add "type","text"
yourname.add "name","yourname"
yourname.add "class","form-control"
yourname.add "maxlength",50
yourname.add "required",true

'##########################################################################################

dim submit : set submit=form.field
submit.add "style","margin-top:10px"
submit.add "type","submit"
submit.add "name","btnAction"
submit.add "value","Submit"
submit.add "class","btn btn-primary"

'##########################################################################################

form.build()

%>