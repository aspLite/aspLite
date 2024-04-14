<%
dim form : set form=aspl.form

if form.postback then
	form.alert "success","Option " & aspl.getRequest("selectbox") & " was selected!"	
end if

dim options : set options=aspl.dict

for i=1 to 1000
	options.add i,"Option " & aspl.padLeft(i,4,0)
next

form.write "Basic Select2 example for aspLite"

dim selectbox : set selectbox=form.field("select")
selectbox.add "class","form-control"
selectbox.add "name","selectbox"
selectbox.add "id","js-data-example-basic"
selectbox.add "options",options
selectbox.add "onchange",form.submit

form.writejs "$('#js-data-example-basic').select2();"

form.build
%>