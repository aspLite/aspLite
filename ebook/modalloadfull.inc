<%
dim form : set form=aspl.form

if form.postback then
	
	form.writejs "aspAjax('GET',aspLiteAjaxHandler,'asplEvent=modalTitle',getAspForm);"
	form.writejs "aspAjax('GET',aspLiteAjaxHandler,'asplEvent=fields&asplTarget=modalBody',getAspForm);"
	form.writejs "aspAjax('GET',aspLiteAjaxHandler,'asplEvent=modalFooter',getAspForm);"
	
	'open the modal
	form.writejs "$('#myModal').modal('show');"
	
end if

dim submit : set submit=form.field("submit")
submit.add "class","btn btn-primary"
submit.add "html","Open modal with a fully featured interactive asplForm"
form.build
%>