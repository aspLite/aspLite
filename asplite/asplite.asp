<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
%>
<!-- #include file="config.asp"-->
<%
dim aspL
set aspL=new cls_asplite

class cls_asplite

	private debug,startTime,stopTime,plugins,p_fso,cacheprefix
	
	Private Sub Class_Initialize()
	
		on error resume next
	
		startTime=Timer()
		debug						= aspL_debug 
		
		'IMPORTANT
		'no matter which language you speak or what you're up to in classic ASP,
		'this is how an ASP page should start (together with the @language, codepage
		'and option explicit statements - see above
		'-------------------------------------------
		Response.Buffer				= true		
		Response.CharSet			= "utf-8" 'does not work on IIS5 (Windows 2000 Servers)
		Response.ContentType		= "text/html"
		Response.CacheControl		= "no-cache"
		Response.AddHeader "pragma", "no-cache"
		Response.Expires			= -1
		Response.ExpiresAbsolute	= Now()-1
		'-------------------------------------------
		
		cacheprefix="asplite_"
		set plugins=nothing	
		set p_fso=nothing
		
		on error goto 0
		
	End Sub	
	
	Private Sub Class_Terminate()
	
		'destroy all plugins
		
		if not plugins is nothing then
			dim p
			for each p in plugins
				set plugins(p)=nothing			
			next
		
			set plugins=nothing
		end if
		
		set p_fso=nothing
		
	End sub
	
	private function removeCRB(value)
	
		'remove Code Render Blocks
	
		value=replace(value,"<" & "%","",1,-1,1)
		removeCRB=replace(value,"%" & ">","",1,-1,1)	
	
	end function	

	'executeGlobal: the ASP code in the file (path) will
	'be executed as classic ASP and it will be available in
	'the namespace of this page-script.
	public sub exec(path)		
		
		on error resume next
		
		executeGlobal removeCRB(stream(path,false,""))
		
		aspError("problem when executing " & path & " - <strong><a target=""_blank"" href=""" & path & """>try to load the ASP file directly</a></strong>")
		
		on error goto 0

	end sub	
	
	'loads the content of text file (utf8)
	public function loadText(path)		
		
		loadText=stream(path,false,"")

	end function
	
	'loads the content of a binary file (image, pdf, zip, etc)
	public function loadBinary(path)		
		
		loadBinary=stream(path,true,"")

	end function	
	
	private function stream(path,binary,byref size)
	
		on error resume next

		Dim objStream
		Set objStream = server.CreateObject("ADODB.Stream")	
		
		if binary then
		
			objStream.Open	
			objStream.type=1 'adTypeBinary
			objStream.LoadFromFile(server.mappath(path))
			stream=objStream.Read()
		
		else
	
			objStream.CharSet = "utf-8"
			objStream.Open	
			objStream.type=2 'adTypeText
			objStream.LoadFromFile(server.mappath(path))
			stream = objStream.ReadText()
			
		end if
		
		size=objStream.size
			
		set objStream=nothing
		
		asperror(path)			
		
		on error goto 0	
	
	end function	
	
	public sub saveText(sPath,data)
	
		on error resume next
	
		Dim objStream
		Set objStream = CreateObject("ADODB.Stream")
		objStream.CharSet = "utf-8"
		objStream.mode = 3 'adModeReadWrite
		objStream.Open		
		objStream.position = 0
		objStream.WriteText data
		objStream.SaveToFile server.mappath(sPath), 2      
		objStream.close
		set objStream = nothing  

		asperror(path)
      
        on error goto 0
		
    end sub	
	
	public function plugin(value)
	
		value=lcase(value)
		
		if plugins is nothing then
			set plugins=dict
		end if
	
		if not plugins.exists(value) then
			
			exec(aspL_path & "/plugins/" & value & "/" & value & ".asp")	
			
			dim pluginCls
			set pluginCls=eval("new cls_asplite_" & value)
			
			plugins.add value,pluginCls
		
		end if
		
		set plugin=plugins(value)
	
	end function	
	
	public function getRequest(value)
	
		on error resume next
	
		if not isEmp(request.querystring(value)) then
			getRequest=request.querystring(value)
		elseif isEmp(request.form(value)) then
			getRequest=request.form(value)
		else
			getRequest=request(value)
		end if
		
		on error goto 0	
	
	end function	
	
	Public function asperror(value)	
	
		if not debug then exit function
		
		if err.number<>0 then
		
			asperror="<h1>aspLite error details:</h1>"
			asperror=asperror & "<span style=""color:Red;font-size:1.5em;font-weight:700"">" & value & "</span><br><br>"
			asperror=asperror & "err.number: " &  err.number & "<br>"
			asperror=asperror & "err.source: " &  err.source & "<br>"
			asperror=asperror & "err.description: " &  err.description
			
			asperror=asperror & "<hr>"
			
			asperror=asperror & replace(request.servervariables("ALL_RAW"),vbcrlf,"<br>",1,-1,1)		
			
			dump asperror
		
		end if
	
	end function	
	
	public sub die
	
		Class_Terminate()		
		response.end	
	
	end sub
	
	public function dump (value)
	
		response.clear
		response.write value
		die()
	
	end function	
	
	public function dumpJson (value)	
		
		Response.ContentType = "application/json"
		dump(value)
	
	end function
	
	public function dumpBinary (path)
	
		on error resume next
	
		path=server.mappath(path)
	
		Dim objStream : Set objStream = server.CreateObject("ADODB.Stream")
		
		objStream.Open	
		objStream.type=1 'adTypeBinary
		objStream.LoadFromFile(path)

		asperror(path)
				
		'get filesize
		dim size : size=objStream.size	

		'set chunksize - files will be served by chunks of 500kb each
		dim chunksize : chunksize=500000
	
		'retrieve filename
		dim filename : filename=right(path,len(path)-InStrRev(path,"\",-1,1))		
		
		'retrieve filetype		
		dim filetype : filetype=getFileType(filename)		
		
		select case lcase(filetype)
		
			case "jpeg","jpg" : response.ContentType="image/JPEG"
			case "png" : response.ContentType="image/x-png"
			case "htm","html" : response.ContentType="text/HTML"
			case "js" : response.ContentType="text/HTML"
			case "gif" : response.ContentType="image/GIF"
			case "txt","css" : response.ContentType="text/plain"
			case "zip" : response.ContentType="application/x-zip-compressed"
			case "pdf" : response.ContentType="application/pdf"
			case "doc","docx" : Response.ContentType = "application/msword"
			case "xls","xlsx" : Response.ContentType = "application/x-msexcel"
			case "mpeg" : Response.ContentType = "video/mpeg"	
			case "mp3" : Response.ContentType = "audio/mpeg"
			case "mp4" : Response.ContentType = "video/mp4"
			case "avi" : Response.ContentType = "video/x-msvideo"
			case "wmv" : Response.ContentType = "video/x-ms-wmv"
			case "m4v" : Response.ContentType = "video/x-m4v"
			case "mov" : Response.ContentType = "video/quicktime"
			case "3gp" : Response.ContentType = "video/3gpp"
			case "xml" : Response.ContentType = "application/xml"
			case "wav" : Response.ContentType = "audio/wav"
			case else Response.ContentType = "application/octet-stream"
		
		end select
				
		response.clear	
		
		Response.AddHeader "Content-Disposition", "attachment; filename=" & filename		
		
		if size<chunksize then
			response.AddHeader "Content-Length", size			
			response.binarywrite objStream.Read()			
		else
		
			dim i
			for i=0 to size step chunksize
				response.binarywrite objStream.Read(chunksize)
				response.flush()
			next
			
		end if		
		
		set objStream=nothing
		
		response.clear
		
		die()
	
	end function
	
	Public Function printTimer() 	  
	   
		stopTime=Timer()
		
		PrintTimer = round((stopTime - startTime) * 1000,0) 'milliseconds
	
	End Function 
	
	public function xmlhttp(url,binary)
	
		on error resume next
		
		dim oxmlhttp
		Set oxmlhttp = server.createobject("MSXML2.ServerXMLHTTP")

		oxmlhttp.open "GET", url
		oxmlhttp.send
		
		if oXMLHTTP.status=200 then
		
			if binary then
				xmlhttp=oxmlhttp.responseBody
			else
				xmlhttp=oxmlhttp.responseText
			end if
		
		else
		
			xmlhttp=oXMLHTTP.status
		
		end if
		
		set oxmlhttp=nothing
		
		if err.number<>0 then	
			asperror(url)
		end if	
		
		on error goto 0

	end function
	
	Public function xmldom(url)
	
		on error resume next
	
		Set xmlDOM = Server.CreateObject("MSXML2.DOMDocument")
		xmlDOM.async = False
		xmlDOM.setProperty "ServerHTTPRequest", True
		xmlDOM.Load(url)
		
		If xmlDOM.parseError.errorCode <> 0 Then
		
			Set xmlDOM = Server.CreateObject("Msxml2.DOMDocument.6.0")
			xmlDOM.async = false
			xmlDOM.setProperty "ServerHTTPRequest", True
			xmlDom.resolveExternals = False
			
			if xmlDOM.Load(url) then
				err.clear
			end if
			
		End If
	
		if err.number<>0 then	
			asperror(url)
		end if	
		
		on error goto 0
	
	end function
	
	'############################
	'### some caching functions
	'############################
	
	public function setcache(name,value)
	
		'cached items always get a timestamp at the moment they're stored

		dim arr(2)
		arr(0)=Timer()
		arr(1)=value		
	
		application(cacheprefix & name)=arr
	
	end function
	
	public function clearcache(name)
	
		application.contents.remove(cacheprefix & name)
	
	end function
	
	public function getCacheT (name, seconds)
		
		'returns the cached content only if it was stored less than x seconds ago
	
		getCacheT=getcache(name)
		
		if not isEmp(getCacheT) then
			if round((Timer - application(cacheprefix & name)(0)),0) > convertNmbr(seconds) then
				getCacheT=""				
			end if
		end if
		

	end function
	
	public function getCache(name)
	
		'returns the cached content, regardless the timestamp
		on error resume next
		
		getCache=application(cacheprefix & name)(1)
		
		if err.number<>0 then getCache=""
		
		on error goto 0
		
	end function
	
	public function clearAllCache
	
		dim el
		for each el in application.contents		
			if left(el,len(cacheprefix))=cacheprefix then
				application.contents.remove(el)
			end if
		next
	
	end function
	
	public function fso
	
		'gives back a filesystemobject (created only once)
	
		if p_fso is nothing then
			set p_fso=server.createobject("scripting.filesystemobject")
		end if
		
		set fso=p_fso
	
	end function
	
	public function dict
	
		'gives back a scripting.dictionary		
		set dict=server.createobject("scripting.dictionary")
	
	end function
	
	
	Public function pathinfo 'get userfriendly url from 404 request if any
		
		dim ufl
		ufl=Request.ServerVariables("query_string")	
				
		if instr(ufl,"?")>0 then
			ufl=left(ufl,instr(ufl,"?"))
		end if			
		
		if not isEmp(ufl) then
		
			ufl=replace(ufl,":80","",1,-1,1)
			ufl=replace(ufl,":443","",1,-1,1)
			ufl=replace(ufl,"http://" & request.servervariables("http_host") & "/" & aspL_path,"",1,-1,1)
			ufl=replace(ufl,"https://" & request.servervariables("http_host") & "/" & aspL_path,"",1,-1,1)
			ufl=replace(ufl,"http://" & request.servervariables("http_host"),"",1,-1,1)
			ufl=replace(ufl,"https://" & request.servervariables("http_host"),"",1,-1,1)
			ufl=replace(ufl,"404;","",1,-1,1)
			ufl=replace(ufl,"aspxerrorpath=/","",1,-1,1)			
			ufl=replace(ufl,":" & request.servervariables("server_port"),"",1,-1,1)		

			dim uflLoop
			uflLoop=true
			
			while uflLoop
				if instr(ufl,"//")>0 then
					ufl=replace(ufl,"//","/",1,-1,1)
				else
					uflLoop=false
				end if
			wend	
			
			uflLoop=true
			
			while uflLoop
				if left(ufl,1)="/" then
					ufl=right(ufl,len(ufl)-1)
				else
					uflLoop=false
				end if
			wend
			
			uflLoop=true

			while uflLoop
				if right(ufl,1)="/" then
					ufl=left(ufl,len(ufl)-1)
				else
					uflLoop=false
				end if
			wend			
			
			ON Error Resume Next		
			
			if not isEmp(ufl) and IsAlphaNumeric(ufl) then
				Response.Status = "200 OK"
				pathinfo=ufl			
			end if			
				
			ON Error Goto 0
			
		end if
		
	end function
	
	
	'#################################################################################################
	'#################################################################################################
	'###### This is it as far as aspLite is concerned. 
	'###### Below you find some generic VBScript functions I often use in ASP projects
	'###### DO NOT REMOVE or CHANGE them. I use some of these functions aspLite (above)
	'###### and/or in some of the plugins I already developed
	'#################################################################################################
	'#################################################################################################
	
	
	public function getFileType(filename)
	
		getFileType=right(filename,len(filename)-InStrRev(filename,".",-1,1))
	
	end function
	
	public function isNumber(byval value)
		
		on error resume next

		if isEmp(value) then
			isNumber=false
		else
			isNumber=isNumeric(value)
		end if
		
		on error goto 0
		
	end function


	public function isEmp(byval value)
	
		on error resume next
		
		isEmp=false
		
		if isNull(value) then
			isEmp=true
		else
			if isEmpty(value) or trim(value)="" then isEmp=true
		end if
		
		on error goto 0
		
	End Function


	public function convertStr(value)

		on error resume next
		
		if not isnull(value) then
			convertStr=cstr(value)
		else
			convertStr=""
		end if
		
		if err.number<>0 then
			convertStr=value
		end if
		
		on error goto 0
		
	End Function

	public function sanitize(sValue)

		if isEmp(sValue) then
			sanitize=""
		else
			sanitize=replace(sValue,"""","&quot;",1,-1,1)
			sanitize=replace(sanitize,"<","&lt;",1,-1,1)
			sanitize=replace(sanitize,">","&gt;",1,-1,1)
		end if
		
	end function

	public function sanitizeJS(sValue)

		sanitizeJS=replace(sValue,"'","\'",1,-1,1)
		
	end function


	public function convertNmbr(value)

		on error resume next

		if isNumber(value) then 
			convertNmbr=cdbl(value)
		else
			convertNmbr=0
		end if
		
		if err.number<>0 then convertNmbr=0
		
		on error goto 0
					
	End Function

	public function convertBool(value)
		
		On Error Resume Next
		
		if isEmp(value) then
			convertBool=false
			exit function
		end if
		if convertNmbr(value)=1 then
			convertBool=true
			exit function
		end if
		if cstr(value)="0" then
			convertBool=false
			exit function
		end if
		if lcase(cstr(value))="true" then
			convertBool=true
			exit function
		end if
		if lcase(cstr(value))="false" then
			convertBool=false
			exit function
		end if
		if value=true or cBool(value) then
			convertBool=true
			exit function
		end if
		
		convertBool=false
		
		On Error Goto 0
		
	End Function

	public function sqli(str)
		if isEmp(str) then
			sqli=""
		else
			sqli=replace(str,"'","''",1,-1,1)
		end if
	end function	

	public function convert2 (byref getal)
		if len(getal)=1 then 
			convert2="0"&getal
		else
			convert2=getal
		end if
	end function	
		
	Function IsAlphaNumeric(byVal str)
	
		If IsNull(str) Then str = ""
		
		Dim ianRegEx
		Set ianRegEx = New RegExp
		ianRegEx.Pattern = "[^a-z0-9\/\_\-\.]"
		ianRegEx.Global = True
		ianRegEx.IgnoreCase = True
		IsAlphaNumeric = (str = ianRegEx.Replace(str,""))
		
		Set ianRegEx=nothing
		
	End Function

	
end class
%>