<%

'formbuilder sample - built on bootstrap css
aspl.exec("code/demo_asp/formbuilder/formbuilder.asp")

dim form : set form=new cls_formbuilder
form.listenTo "e","sampleform14"
form.targetDiv="sampleform14"
form.id="uploadForm"
form.onSubmit="" 'overrule the default submit - 
'We need to overrule the default submit.
'this is needed here because this multi-fileuploader does not really submit the form
'it sends a collection of files
	
'result
dim feedback : set feedback=form.field
feedback.add "type","comment"
'this is one of a kind... The complete form is included as a comment, but that "comment" includes all the
'html and javascript needed for this jQuery uploader.... Kind of weird, but it works like a charm.
feedback.add "html",aspL.loadText("html/demo_asp/uploadjquery.resx")
feedback.add "tag","div"

form.build()

%>