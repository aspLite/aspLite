<%
dim form, field

select case lcase(aspL.getRequest("asplEvent"))	
	
	case "basicdiv1"
	
		'this aspLite form binds to data-asplTarget="basicDiv1"
		set form=aspl.form
		set field = form.field("plain")
		field.add "html","<p><i>Onload event!</i></p>"
		
		form.build()

	case "clicklink"
	
		'this aspLite form binds to data-asplTarget="basicDiv2"
		set form=aspl.form
		
		set field = form.field("script")
		field.add "text","console.log('you can also send some JavaScript back to the browser')"	
		
		set field = form.field("element")
		field.add "tag","p"
		field.add "html","It's a start! Check the console (F12)"
		
		form.build()		

end select

response.write aspL.loadText("basic/html/basic.resx")

set aspL=nothing
%>