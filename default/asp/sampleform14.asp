<%

dim form : set form=aspl.form
form.target="sampleform14"

form.onSubmit="" '
'We need to overrule the default submit.
'this is needed here because this multi-fileuploader does not really submit the form
'it sends a collection of files with sequental Ajax calls by itself
	
'result
dim feedback : set feedback=form.field("element")
'this is one of a kind... The complete form is included as div-element, but that "element" includes all the
'html and javascript needed for this jQuery uploader.... Kind of weird, but it works like a charm.
feedback.add "html",aspL.loadText("default/html/sampleform14.resx")
feedback.add "tag","div"

form.build()

%>