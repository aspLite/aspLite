<%
select case lcase(aspl.pathinfo)

	case "customurl" : aspL.exec("code/asp/404/default.asp")
	
	case "learn-more" : aspL.exec("code/asp/404/learnmore.asp")	
	
	case "about" : aspL.exec("code/asp/404/about.asp")	
	
	case "getcode" : aspL.exec("code/asp/404/getcode.asp")	
	
	case "contact" : aspL.exec("code/asp/404/contact.asp")	
	
	case else 'do nothing?
	
end select
%>