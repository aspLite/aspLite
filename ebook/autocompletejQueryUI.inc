<%
set form = aspl.form

set field = form.field ("text")
field.add "id","jqAutocomplete"
field.add "class","form-control"
field.add "placeholder","start writing a name"

form.writejs aspl.loadText("ebook/autocompletejQueryUIjs.inc")

form.build
%>