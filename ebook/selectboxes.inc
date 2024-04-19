<%
dim form : set form=aspl.form

'VBScript dictionary
dim i, dictionary : set dictionary=aspl.dict
for i=1 to 9
	dictionary.add i,"option" & i
next

if form.postback then

	dim feedback : feedback="<div class=""alert alert-success"">"
	feedback=feedback & "<p>Selectbox value: <strong>"
	feedback=feedback & aspl.htmlencode(aspl.convertNmbr(aspl.getRequest("selectbox"))) & "</strong></p>"
	feedback=feedback & "<p>Radio value: <strong>"
	feedback=feedback & aspl.htmlencode(aspl.convertNmbr(aspl.getRequest("radio")))& "</strong></p>"
	feedback=feedback & "<p>Checkbox values: <strong>"
	feedback=feedback & aspl.htmlencode(aspl.getRequest("checkboxes"))& "</strong></p>"	
	form.write feedback & "</div>"

end if

dim selectbox : set selectbox=form.field("select")
selectbox.add "name","selectbox"
selectbox.add "emptyfirst",""
selectbox.add "class","form-control form-select"
selectbox.add "label","Selectbox showing all values of a VBScript dictionary"
selectbox.add "options",dictionary
selectbox.add "onchange",form.submit

form.newline

dim radio : set radio=form.field("radio")
radio.add "name","radio"
radio.add "class","form-check-input"
radio.add "labelclass","form-check-label"
radio.add "label","Radiobutton showing all values of a VBScript dictionary"
radio.add "options",dictionary
radio.add "onchange",form.submit

form.newline

form.write "Checkboxes showing all values of a VBScript dictionary"

dim checkboxes : set checkboxes=form.field("checkbox")
checkboxes.add "class","form-check-input"
checkboxes.add "labelclass","form-check-label"
checkboxes.add "name","checkboxes"
checkboxes.add "options",dictionary
checkboxes.add "container","div"
checkboxes.add "containerclass","form-check form-switch"
checkboxes.add "onclick",form.submit

form.build

%>