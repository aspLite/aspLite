<%
'load jquery date-functions (sets the jQuery UI dateformat!)
aspL("default/asp/includes/jQueryUiFunctions.asp")

dim form : set form=aspl.form

'form-submitted
if form.postback then

	form.doScroll=true
	
	'feedback
	dim feedback : set feedback=form.field("element")
	
	dim par
	par="<p>jQuery date:<br><strong>" & FormatDateTime(dateFromPicker(aspl.getRequest("today")),1) & "</strong> <small>(vbLongDate)</small></p>"
	par=par & "<p>HTML date:<br><strong>" & FormatDateTime(aspl.getRequest("datefield"),1) & "</strong> <small>(vbLongDate)</small></p>"
	
	feedback.add "html",par
	feedback.add "tag","div"
	feedback.add "class","alert alert-success"

end if

dim today : set today=form.field("text")
today.add "label","Select a jQuery UI date"
today.add "name","today"
today.add "class","form-control"
today.add "id","today"
today.add "value",dateToPicker(date)

'Does not work on SAFARI for God knows what reason!! (Apple did not like that great HTML5-idea I guess)
dim datefield : set datefield=form.field("date")
datefield.add "label","as opposed to a regular HTML date-field (does not work in Safari)"
datefield.add "name","datefield"
datefield.add "id","datefield"
datefield.add "class","form-control"
datefield.add "value",aspl.convertHtmlDate(date)
datefield.add "onchange","$('#" & form.id & "').submit()"

dim script : set script=form.field("script")
script.add "text","$('#today').datepicker({inline: true, dateFormat:'" & dateformat & "',onSelect: function(){$('#" & form.id & "').submit()}})"

form.build()
%>