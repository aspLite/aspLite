<%
'formbuilder sample - built on bootstrap css
'load some variables and functions (dateformat, dateFromPicker, dateToPicker)
aspL.exec("code/demo_asp/functions.asp")
aspl.exec("code/demo_asp/formbuilder/formbuilder.asp")

dim form : set form=new cls_formbuilder
form.listenTo "e","jqueryui"
form.targetDiv="body"
form.id="dateForm"

dim intro : set intro=form.field
intro.add "type","comment"
intro.add "tag","p"
intro.add "html","Very useful Datepicker by jQuery UI - <a href=""https://jqueryui.com/datepicker/"" target=""_blank""> https://jqueryui.com/datepicker/</a> with lots of options. Check code/demo_asp/jqueryui.asp for the ASP code involved."

'form-submitted
if form.postback then
	
	'feedback
	dim feedback : set feedback=form.field
	feedback.add "html","<p>You selected <strong>" & FormatDateTime(dateFromPicker(aspl.getRequest("dp")),1) & "</strong> (vbLongDate)</p>"
	feedback.add "type","comment"
	feedback.add "tag","div"	
	
	'you can build it here already. This will stop further exection	
	form.build()

end if


dim today : set today=form.field
today.add "label","Select a date"
today.add "type","text"
today.add "name","dp"
today.add "class","form-control"
today.add "id","dp"
today.add "value",dateToPicker(date)
'today.add "onchange","$('#" & form.id & "').submit()" -> would also work

dim script : set script=form.field
script.add "type","script"
script.add "text","$('#dp').datepicker({inline: true, dateFormat:'" & dateformat & "',onSelect: function(){$('#" & form.id & "').submit()}})"

form.build()
%>