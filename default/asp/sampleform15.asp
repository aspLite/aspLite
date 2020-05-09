<%

dim form : set form=aspl.form
form.listenTo "asplEvent","sampleform15"
form.target="sampleform15"

dim defaultValue
defaultValue=aspl.loadText("default/asp/includes/codemirrorsample.asp")
	
'form-submitted
if form.postback then

	defaultValue=aspl.getRequest("cm1")
	
	'feedback
	dim feedback : set feedback=form.field("element")
	feedback.add "html","Saved!"
	feedback.add "tag","div"
	feedback.add "class","alert alert-info"		
	
	'you can build it here already. This will stop further exection	
	'form.build()

end if	

'##########################################################################################

dim cm1 : set cm1=form.field("textarea")
cm1.add "name","cm1"
cm1.add "id","cm1"
cm1.add "value",defaultValue

dim scriptSrc : set scriptSrc=form.field("script")
scriptSrc.add "text", "var editor;editor = CodeMirror.fromTextArea(document.getElementById('cm1') , { lineNumbers: true,isASP: true})"

'##########################################################################################

dim submit : set submit=form.field("submit")
submit.add "style","margin-top:10px;margin-bottom:10px"
submit.add "name","btnAction"
submit.add "value","Save"
submit.add "class","btn btn-primary"
submit.add "onclick","editor.save()"

'##########################################################################################

form.build()

%>