<%
'formbuilder sample - built on bootstrap css
'load some variables and functions (dateformat, dateFromPicker, dateToPicker)
aspL.exec("code/demo_asp/functions.asp")
aspl.exec("code/demo_asp/formbuilder/formbuilder.asp")

dim form : set form=new cls_formbuilder
form.targetDiv="body"
form.requiredLegend=""
form.requiredStar=""
form.id="dateForm"

dim intro : set intro=aspl.dict
intro.add "type","comment"
intro.add "tag","div"
intro.add "html",aspL.loadText("html/demo_asp/jqueryui.resx")
form.addField(intro)

'form-submitted
if form.postback then
	
	'feedback
	dim feedback : set feedback=aspl.dict
	feedback.add "html","<p>You selected <strong>" & FormatDateTime(dateFromPicker(aspl.getRequest("dp")),1) & "</strong> (vbLongDate)</p>"
	feedback.add "type","comment"
	feedback.add "tag","div"
	form.addField(feedback)	
	
	'you can build it here already. This will stop further exection	
	form.build()

end if

'this hidden field is required in this demo, as the "e"-event is used in the event-handler
dim hidden : set hidden=aspl.dict
hidden.add "type","hidden"
hidden.add "name","e"
hidden.add "value","jqueryui"

form.addField(hidden)

dim today : set today=aspl.dict
today.add "label","Select a date"
today.add "type","text"
today.add "name","dp"
today.add "class","form-control"
today.add "id","dp"
today.add "value",dateToPicker(date)
'today.add "onchange","$('#" & form.id & "').submit()" -> would also work

form.addField(today)

dim script : set script=aspl.dict
script.add "type","comment"
script.add "tag","script"
script.add "html","$('#dp').datepicker({inline: true, dateFormat:'" & dateformat & "',onSelect: function(){$('#" & form.id & "').submit()}})"
	
form.addField(script)

form.build()
%>