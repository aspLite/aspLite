<%
'login/logout form, with 3 attempts allowed
set form=aspl.form

if form.postback then

	set script=form.field("script")
	script.add "text","$.getScript('https://unpkg.com/xlsx@0.16.0/dist/xlsx.core.min.js',function(){createXLS()})"
	
	set script=form.field("script")
	script.add "text",aspl.loadText("default/html/sampleform30.resx")

end if

set field = form.field("submit")
field.add "value","click to generate xlsx"
field.add "class","btn btn-success"

form.build

%>