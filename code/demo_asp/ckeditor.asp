<%
'formbuilder sample - built on bootstrap css
aspl.exec("code/demo_asp/formbuilder/formbuilder.asp")

dim form : set form=new cls_formbuilder
form.listenTo "e","ckeditor"
form.targetDiv="body"
form.id="ckeditor4"

dim intro : set intro=form.field
intro.add "type","comment"
intro.add "html","<p>CKEditor 4 - <a target=""_blank"" href=""https://ckeditor.com/ckeditor-4/"">https://ckeditor.com/ckeditor-4/</a></p>"
intro.add "tag","div"
	
'form-submitted
if form.postback then	
	
	'feedback
	dim feedback : set feedback=form.field
	feedback.add "type","comment"
	feedback.add "html","Saved!"
	feedback.add "tag","div"
	feedback.add "class","alert alert-info"		
	
	'you can build it here already. This will stop further exection	
	'form.build()

end if	

'##########################################################################################

dim mNotes1 : set mNotes1=form.field
mNotes1.add "type","textarea"
mNotes1.add "name","mNotes1"
mNotes1.add "id","mNotes1"
mNotes1.add "value",aspl.getRequest("mNotes1")

dim mNotes2 : set mNotes2=form.field
mNotes2.add "type","textarea"
mNotes2.add "name","mNotes2"
mNotes2.add "id","mNotes2"
mNotes2.add "value",aspl.getRequest("mNotes2")

dim scriptSrc : set scriptSrc=form.field
scriptSrc.add "type","script"
scriptSrc.add "text", "$.getScript( '//cdn.ckeditor.com/4.14.0/basic/ckeditor.js',function(){CKEDITOR.replace( 'mNotes1' );CKEDITOR.replace( 'mNotes2' )})"

'##########################################################################################

dim submit : set submit=form.field
submit.add "style","margin-top:10px"
submit.add "type","submit"
submit.add "name","btnAction"
submit.add "id","btnAction"
submit.add "value","Save"
submit.add "class","btn btn-primary"
submit.add "onclick","CKEDITOR.instances.mNotes1.updateElement();CKEDITOR.instances.mNotes2.updateElement()"

'##########################################################################################

form.build()

%>