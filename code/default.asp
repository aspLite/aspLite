<%
option explicit  'include this line in all your code files - it ensures that you declare all variables you'll use.
asp.codebehind() 'include this line in all your code files - it ensures that this page does not load when called directly.

dim titletag,body

titletag="Sample page"

'here you can typically add some sort of eventhandler

select case lcase(asp.getRequest("action"))
		
	case "clicklink"
		body="<p>Link was clicked</p>"
		'just for the fun of it... add some JavaScript in the header
		asp.js.addHEAD "window.onload=function(e){document.body.style.backgroundColor='#FAA'}"
		
		
	case "clickbutton"
		body="<p>Button was clicked</p>"
		asp.js.addHEAD "window.onload=function(e){document.body.style.backgroundColor='#FF0'}"

		
		'in this case, add some javascript at the bottom too
		asp.js.addBODY "document.write ('JavaScript added right before the closing body-tag.')"
		
		
	case "loadclass"
	
		'CONDITIONAL load of an asp page/class
		asp.ASP_executeGlobal("code/class.asp")
		
		asp.js.addHEAD "window.onload=function(e){document.body.style.backgroundColor='#F5F'}"
		
		dim testObj
		set testObj=new cls_test
		body="<p>" & testObj.hello & "</p>"
		set testObj=nothing
		
	case else
	
		body="<p>No (known) action was detected. Initial load.</p>"
		
		'add some JavaScript code in the header of the page
		asp.js.addHEAD "window.onload=function(e){document.body.style.backgroundColor='#FFF'}"

end select


%>