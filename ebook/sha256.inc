<%
dim form : set form=aspl.form

dim password : password=aspl.randomizer.createGUID(10)

form.write "<strong>" & password & "</strong><br>"
form.write " hashes to <br>"
form.write " <strong>" & aspL.plugin("sha256").sha256(password) & "</strong>"

form.newline

dim reload : set reload=form.field("submit")
reload.add "html","Reload"
reload.add "class","btn btn-danger"

form.build
%>