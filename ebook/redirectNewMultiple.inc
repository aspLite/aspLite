<%

dim form : set form=aspl.form
form.alert "info",aspl.htmlEncode(aspl.getRequest("message"))
form.build

%>