<%

dim form : set form=aspl.form
form.target="sampleform0"

dim feedback : set feedback=form.field("element")
feedback.add "tag","div"

feedback.add "class","alert alert-success"
feedback.add "html","This is a demo SPA (Single Page Application) making use of <strong>aspLite</strong>. Various sample forms, plugins and utilities are showcased over here. The location of the ASP-code involved, is displayed each time. Have fun while discovering the bits and bytes of this demo!"

form.build()

%>