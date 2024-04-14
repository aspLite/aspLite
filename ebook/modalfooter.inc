<%
dim form : set form=aspl.form

set close=form.field("button")
close.add "html","Close"
close.add "class","btn btn-secondary"
close.add "databsdismiss","modal"

form.build
%>