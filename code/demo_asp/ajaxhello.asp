<%
'formbuilder sample - built on bootstrap css
aspl.exec("code/demo_asp/formbuilder/formbuilder.asp")

dim form : set form=new cls_formbuilder
form.listenTo "e","ajaxhello"
form.targetDiv="ajaxhello"
form.id="helloForm"
form.requiredLegend=""
form.requiredStar=""
form.doScroll=false

'form-submitted
if form.postback then

	form.doScroll=true
	
	'feedback
	dim feedback : set feedback=aspl.dict
	feedback.add "type","comment"
	feedback.add "html","Hello " & aspL.sanitize(aspL.getRequest("yourname")) & "! Check /code/demo_asp/ajaxhello.asp!"
	feedback.add "tag","div"
	feedback.add "class","alert alert-warning"	
	form.addField(feedback)	
	
	'you can build it here already. This will stop further exection	
	form.build()

end if	

'##########################################################################################

dim yourname : set yourname=aspl.dict
yourname.add "placeholder","Your name:"
yourname.add "type","text"
yourname.add "name","yourname"
yourname.add "id","yourname"
yourname.add "class","form-control"
yourname.add "maxlength",50
yourname.add "required",true

form.addField(yourname)

'##########################################################################################

dim submit : set submit=aspl.dict
submit.add "style","margin-top:10px"
submit.add "type","submit"
submit.add "name","btnAction"
submit.add "value","Submit"
submit.add "class","btn btn-primary"

form.addField(submit)

'##########################################################################################

form.build()

%>