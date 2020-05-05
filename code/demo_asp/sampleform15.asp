<%

dim form : set form=new cls_formbuilder
form.listenTo "e","sampleform15"
form.targetDiv="sampleform15"
form.id="codeMirrirEditor"

dim style : set style=form.field
style.add "type","link"
style.add "href","https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.52.2/codemirror.min.css"
style.add "rel","stylesheet"

dim inlineStyle : set inlineStyle=form.field
inlineStyle.add "type","style"
inlineStyle.add "text",".CodeMirror { border: 1px solid #eee;height:auto}"

dim defaultValue
defaultValue=aspl.loadText("code/demo_asp/codemirrorsample.asp")
	
'form-submitted
if form.postback then

	defaultValue=aspl.getRequest("cm1")
	
	'feedback
	dim feedback : set feedback=form.field
	feedback.add "type","element"
	feedback.add "html","Saved!"
	feedback.add "tag","div"
	feedback.add "class","alert alert-info"		
	
	'you can build it here already. This will stop further exection	
	'form.build()

end if	

'##########################################################################################

dim cm1 : set cm1=form.field
cm1.add "type","textarea"
cm1.add "name","cm1"
cm1.add "id","cm1"
cm1.add "value",defaultValue

dim scriptSrc : set scriptSrc=form.field
scriptSrc.add "type","script"
scriptSrc.add "text", "var editor;$.getScript('https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.52.2/codemirror.min.js', function() {$.getScript('https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.52.2/mode/vbscript/vbscript.min.js',function(){editor = CodeMirror.fromTextArea(document.getElementById('cm1') , { lineNumbers: true,isASP: true})})})"

'##########################################################################################

dim submit : set submit=form.field
submit.add "style","margin-top:10px;margin-bottom:10px"
submit.add "type","submit"
submit.add "name","btnAction"
submit.add "value","Save"
submit.add "class","btn btn-primary"
submit.add "onclick","editor.save()"

'##########################################################################################

form.build()

%>