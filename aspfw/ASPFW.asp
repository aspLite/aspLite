<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit 'include this line in all your code files - it ensures that you declare all variables you use.

'this class is the heart of the framework
class cls_aspfw	

	Private Sub Class_Initialize()
	
		'this is how ASP applications start. 
		'These lines deal with buffering, charset UTF-8, prevent caching, and set the correct contenttype 
	
		Response.Buffer				= true
		session.Timeout				= 180 '3 hours
		server.ScriptTimeout		= 800 '3 minutes: needed for uploading bigger pictures/files!
		Response.CharSet			= "utf-8"
		Response.ContentType		= "text/html"
		Response.CacheControl		= "no-cache"
		Response.AddHeader "pragma", "no-cache"
		Response.Expires			= -1
		Response.ExpiresAbsolute	= Now()-1
	
		'the function ASP_executeGlobal is identical to server #include directives.
		'however, you can CONDITIONALLY include these ASP files. that is how you can keep RAM usage low(er)
		'resx files are never served to browsers, so they are safer to use
		ASP_executeGlobal("aspfw/functions.resx")
		ASP_executeGlobal("aspfw/timer.resx")
		ASP_executeGlobal("aspfw/javascript.resx")		
		ASP_executeGlobal("aspfw/messages.resx")
		
	End Sub	

	public sub ASP_executeGlobal(path)		

		path=lcase(path)
		
		dim strData
		strData=ASP_loadfile(path)
		
		strData=replace(strData,"<" & "%","",1,-1,1)
		strData=replace(strData,"%" & ">","",1,-1,1)	
		
		executeGlobal strData	

	end sub	
	
	public function ASP_loadfile(path)

		Dim objStream
		Set objStream = server.CreateObject("ADODB.Stream")
		objStream.CharSet = "utf-8"
		objStream.Open	
		objStream.LoadFromFile(server.mappath(path))
		ASP_loadfile = objStream.ReadText()	
		set objStream=nothing

	end function

end class

%>