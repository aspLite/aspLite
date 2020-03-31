<!-- #include file="asp/ASP.asp"-->
<%

select case lcase(asp.getrequest("myaction"))

	case "onload"		
	
		asp.flush "Hello world, Καλημέρα κόσμε (utf8-ready)"

	case "submit1"
	
		asp.flush "Form 1 save button was clicked - " & asp.getrequest("yourname")
		
	case "linksubmit"
	
		asp.flush "Form 1 submitted by link - " & asp.getrequest("yourname")
	
	case "delete1"
	
		asp.flush "Form 1 delete button was clicked - " & asp.getrequest("yourname")
		
	case "submit2"
	
		asp.flush "Form 2 save button was clicked"
	
	case "delete2"
	
		asp.flush "Form 2 delete button was clicked"
		
	case "buttonclick"
	
		asp.flush "Regular button was clicked"
		
	case "clicklink1"

		asp.flush "First link clicked..."
	
	case "clicklink2"

		asp.flush "Second link clicked..."	
	
	case "returnbool"
	
		'returns a random boolean to the browser		
		asp.flush asp.plugin("randomizer").randomnumber(0,100)>49	
		
	case else 'initial pageload!
	
		asp.flush asp.ASP_loadfile("html/ajax.resx")

end select

%>