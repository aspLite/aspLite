<%
'formbuilder sample - built on bootstrap css

aspl.exec("code/demo_asp/formbuilder/formbuilder.asp")
dim form : set form=new cls_formbuilder
form.targetDiv="ajaxhello"
form.requiredLegend=""
form.requiredStar=""
form.doScroll=false

'form-submitted
if form.postback then

	form.doScroll=true
	
	'feedback
	dim feedback : set feedback=aspl.dict
	feedback.add "html","Hello " & aspL.sanitize(aspL.getRequest("yourname")) & "! Check /code/demo_asp/ajaxhello.asp!"
	feedback.add "type","comment"
	feedback.add "tag","div"
	feedback.add "class","alert alert-warning"	
	form.addField(feedback)	
	
	'you can build it here already. This will stop further exection	
	form.build()

end if	

'this hidden field is required in this demo, as the "e"-event is used in the event-handler
dim hidden : set hidden=aspl.dict
hidden.add "type","hidden"
hidden.add "name","e"
hidden.add "value","ajaxhello"

form.addField(hidden)

'##########################################################################################

dim yourname : set yourname=aspl.dict
yourname.add "placeholder","Your name:"
yourname.add "type","text"
yourname.add "name","yourname"
yourname.add "id","yname"
yourname.add "class","form-control"
yourname.add "maxlength",50
yourname.add "required",true

form.addField(yourname)

'##########################################################################################

dim submit : set submit=aspl.dict
submit.add "style","margin-top:10px"
submit.add "type","submit"
submit.add "name","btnAction"
submit.add "id","btnAction"
submit.add "value","Submit"
submit.add "class","btn btn-primary"

form.addField(submit)

'##########################################################################################

form.build()

%>