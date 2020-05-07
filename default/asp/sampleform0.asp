<%

dim form : set form=aspl.form
form.target="body"

dim feedback : set feedback=form.field("plain")
feedback.add "html",aspl.loadText("default/html/sampleform0.resx")

form.build()

%>