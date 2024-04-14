<%
dim form : set form=aspl.form

if form.postback then
	form.alert "success","Id: " & aspl.getRequest("selectbox") & " was selected!"
	form.build
end if

form.write "AJAX Select2 example for aspLite"

dim selectbox : set selectbox=form.field("select")
selectbox.add "class","form-control"
selectbox.add "name","selectbox"
selectbox.add "id","js-data-example-ajax"

form.writejs aspl.loadText("ebook/select2js.inc")

form.newline

dim submit : set submit=form.field("submit")
submit.add "html","Submit"
submit.add "class","btn btn-primary"

form.build
%>