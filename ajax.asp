<!-- #include file="asp/ASP.asp"-->
<%

select case lcase(asp.getrequest("myaction"))

	case "onload"		
	
		asp.flush "<p>Hello world, Καλημέρα κόσμε (utf8-ready)</p>"

	case "submit1"
	
		asp.flush "<p>Form 1 save button was clicked</p>"
	
	case "delete1"
	
		asp.flush "<p>Form 1 delete button was clicked</p>"
		
	case "submit2"
	
		asp.flush "<p>Form 2 save button was clicked</p>"
	
	case "delete2"
	
		asp.flush "<p>Form 2 delete button was clicked</p>"
		
	case "buttonclick"
	
		asp.flush "<p>Regular button was clicked</p>"
		
	case "clicklink1"

		asp.flush "<p>First link clicked...</p>"
	
	case "clicklink2"

		asp.flush "<p>Second link clicked...</p>"
		

	case else 'initial pageload!
	
		asp.flush asp.ASP_loadfile("html/ajax.resx")

end select

%>