<%
select case lcase(aspl.pathinfo)

	case "customurl" 	: aspL("default/asp/404/default.asp")
	
	case "learn-more" 	: aspL("default/asp/404/learnmore.asp")	
	
	case "about" 		: aspL("default/asp/404/about.asp")	
	
	case "getcode" 		: aspL("default/asp/404/getcode.asp")	
	
	case "contact" 		: aspL("default/asp/404/contact.asp")
	
	case "" 'do nothing
	
	case else 
		
		response.write "this page does not exist"
		response.end
	
end select
%>