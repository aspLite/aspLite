<%
dim form : set form=aspl.form

dim autocomplete : set autocomplete=form.field("text")
autocomplete.add "name","autocomplete"
autocomplete.add "id","autocompleteJQ"
autocomplete.add "placeholder","start writing a name"
autocomplete.add "class","form-control"

form.writejs aspl.loadText("ebook/autocompletejs.inc")

form.build
%>