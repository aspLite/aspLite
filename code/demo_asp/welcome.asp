<%

dim form : set form=new cls_formbuilder
form.targetDiv="welcome"
form.id="welcomeText"

dim feedback : set feedback=form.field
feedback.add "type","element" 
feedback.add "tag","div"
feedback.add "style","margin-top:10px;margin-bottom:20px"

feedback.add "class","alert alert-success"
feedback.add "html","This is a demo SPA (Single Page Application) making use of <strong>aspLite</strong>. Various sample forms, plugins and utilities are showcased over here. The location of the ASP-code involved, is displayed each time. Have fun while discovering the bits and bytes of this demo!"

form.build()

%>