<%
'load some variables and functions (dateformat, dateFromPicker, dateToPicker)
aspL.exec("code/demo_asp/functions.asp")

dim form : set form=new cls_formbuilder
form.listenTo "e","sampleform3"
form.targetDiv="sampleform3"
form.id="dateForm"

dim intro : set intro=form.field
intro.add "type","comment"
intro.add "tag","p"
intro.add "html","<a href=""https://jqueryui.com/datepicker/"" target=""_blank""> https://jqueryui.com/datepicker/</a>"

'form-submitted
if form.postback then

	form.doScroll=true
	
	'feedback
	dim feedback : set feedback=form.field
	feedback.add "html","<p>You selected <strong>" & FormatDateTime(dateFromPicker(aspl.getRequest("dp")),1) & "</strong> (vbLongDate)</p>"
	feedback.add "type","comment"
	feedback.add "tag","div"
	feedback.add "class","alert alert-success"

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