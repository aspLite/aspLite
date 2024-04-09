<%

dim form : set form=aspl.form
form.alert "info",aspl.htmlEncode(aspl.getRequest("message"))

dim back : set back=form.field("button")
back.add "html","Back"
back.add "class","btn btn-secondary"
back.add "onclick",form.redirect("redirect","","")

form.build

%>