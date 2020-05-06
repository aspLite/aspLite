<%
'load some variables and functions (dateformat, dateFromPicker, dateToPicker)
aspL.exec("code/asp/functions.asp")

dim form : set form=aspl.form
form.listenTo "e","sampleform3"
form.target="sampleform3"

'form-submitted
if form.postback then

	form.doScroll=true
	
	'feedback
	dim feedback : set feedback=form.field("element")
	feedback.add "html","<p>You selected <strong>" & FormatDateTime(dateFromPicker(aspl.getRequest("today")),1) & "</strong> (vbLongDate)</p>"
	feedback.add "tag","div"
	feedback.add "class","alert alert-success"

end if

dim today : set today=form.field("text")
today.add "label","Select a date"
today.add "name","today"
today.add "class","form-control"
today.add "id","today"
today.add "value",dateToPicker(date)
'today.add "onchange","$('#" & form.id & "').submit()" -> would also work

dim script : set script=form.field("script")
script.add "text","$('#today').datepicker({inline: true, dateFormat:'" & dateformat & "',onSelect: function(){$('#" & form.id & "').submit()}})"

form.build()
%>