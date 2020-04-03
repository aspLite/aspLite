<%
'###########################################################################
'## include these 2 lines to start each codebehind file for security reasons
option explicit  
asp.codebehind() 
'###########################################################################

dim html,titletag,body

'load the template file
'resx files are never served to browsers, so they are safer to use
html=asp.load("html/default.resx")

titletag="aspLite"

'here you can typically add some sort of eventhandler (what exactly was clicked on?)
select case lcase(asp.getRequest("action"))
		
	case "clicklink"
		
		body="<p>Link was clicked</p>"		
		
	case "clickbutton"
	
		body="<p>Button was clicked</p>"		
		
	case "loadclass"
	
		'CONDITIONAL load of asp page = include file
		asp.exec("code/includes/class.asp")	
		
		dim testObj
		set testObj=new cls_test
		body="<p>" & testObj.hello & "</p>"
		set testObj=nothing	
		
	case "sendhelloajax"		

		asp.flush "Hello " & asp.sanitize(asp.URLDecode(asp.getRequest("yourname")))
	
	case "downloadlargefile"
	
		asp.flushBinaryFile("html/largefile.jpg")
	
	case "downloadsmallfile"
	
		asp.flushBinaryFile("html/smallfile.jpg")
		
	case else
	
		body="<p>No (known) action was detected. Initial load.</p>"		

end select



%>