<%

dim form : set form=aspl.form
form.target="sampleform14"

form.onSubmit="" '
'We need to overrule the default submit.
'this is needed here because this multi-fileuploader does not really submit the form
'it sends a collection of files with sequental Ajax calls by itself
	
'result
dim feedback : set feedback=form.field("plain")
feedback.add "html",aspL.loadText("default/html/sampleform14.resx")

form.build()

%>