<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
%>
<!-- #include file="config.asp"-->
<%
dim aspL
set aspL=new cls_asplite

class cls_asplite	

	private debug,startTime,stopTime,plugins,p_fso
	
	Private Sub Class_Initialize()
	
		on error resume next
	
		startTime=Timer()
		debug						= true 'you may want to set to false for live sites	
		
		'IMPORTANT
		'no matter what language you speak or what you're up to in classic ASP,
		'this is how an ASP page should start (together with the language, codepage
		'and option explicit statement - see above
		'-------------------------------------------
		Response.Buffer				= true		
		Response.CharSet			= "utf-8" 'does not work on IIS5 (Windows 2000 Servers)
		Response.ContentType		= "text/html"
		Response.CacheControl		= "no-cache"
		Response.AddHeader "pragma", "no-cache"
		Response.Expires			= -1
		Response.ExpiresAbsolute	= Now()-1
		'-------------------------------------------
		
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
		
		aspError("problem when executing " & path)
		
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
			set plugins=server.createobject("scripting.dictionary")
		end if
	
		if not plugins.exists(value) then
			
			exec(asp_path & "/plugins/" & value & "/" & value & ".resx")	
			
			dim pluginCls
			set pluginCls=eval("new cls_asplite_" & value)
			
			plugins.add value,pluginCls
		
		end if
		
		set plugin=plugins(value)
	
	end function	
	
	public function getRequest(value)
	
		on error resume next
	
		if not isLeeg(request.querystring(value)) then
			getRequest=request.querystring(value)
		elseif isLeeg(request.form(value)) then
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
			asperror=asperror & value & "<br><br>"
			asperror=asperror & "err.number: " &  err.number & "<br><br>"
			asperror=asperror & "err.description: " &  err.description & "<br><br>"
			
			asperror=asperror & request.servervariables("ALL_RAW")			
			
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
	
		application("aspLite_" & name)=value
	
	end function
	
	public function clearcache(name)
	
		application.contents.remove("aspLite_" & name)
	
	end function
	
	public function getcache(name)
	
		getcache=application("aspLite_" & name)
		
	end function
	
	public function clearAllCache
	
		dim el
		for each el in application.contents		
			if left(el,4)="aspLite_" then
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
	
	
	Public function pathinfo 'get userfriendly url from 404 request if any
		
		dim ufl
		ufl=Request.ServerVariables("query_string")	
				
		if instr(ufl,"?")>0 then
			ufl=left(ufl,instr(ufl,"?"))
		end if			
		
		if not isLeeg(ufl) then
		
			ufl=replace(ufl,":80","",1,-1,1)
			ufl=replace(ufl,":443","",1,-1,1)
			ufl=replace(ufl,"http://" & request.servervariables("http_host") & "/" & asp_path,"",1,-1,1)
			ufl=replace(ufl,"https://" & request.servervariables("http_host") & "/" & asp_path,"",1,-1,1)
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
			
			if not isLeeg(ufl) and IsAlphaNumeric(ufl) then
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

		if isLeeg(value) then
			isNumber=false
		else
			isNumber=isNumeric(value)
		end if
		
	end function


	public function isLeeg(byval value)
		
		isLeeg=false
		
		if isNull(value) then
			isLeeg=true
		else
			if isEmpty(value) or trim(value)="" then isLeeg=true
		end if
		
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

		if isLeeg(sValue) then
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


	public function convertGetal(value)

		on error resume next

		if isNumber(value) then 
			convertGetal=cdbl(value)
		else
			convertGetal=0
		end if
		
		if err.number<>0 then convertGetal=0
		
		on error goto 0
					
	End Function

	public function convertBool(value)
		
		On Error Resume Next
		
		if isLeeg(value) then
			convertBool=false
			exit function
		end if
		if convertGetal(value)=1 then
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
		if isLeeg(str) then
			sqli=""
		else
			sqli=replace(str,"'","''",1,-1,1)
		end if
	end function


	public function numberList(startNr,stopNr,interval,selected)
		dim i
		for i=startNr to stopNr step interval
			numberList=numberList& "<option value=" & i 
			if convertGetal(selected)=convertGetal(i) then numberList=numberList & " selected"
			numberList=numberList& ">" & i & "</option>"
		next
	end function		

	public function convert2 (byref getal)
		if len(getal)=1 then 
			convert2="0"&getal
		else
			convert2=getal
		end if
	end function
	

	Public Function URLDecode(value)	
		
		If isLeeg(value) Then
		   URLDecode = ""
		   Exit Function
		End If
		
		dim aSplit, sOutput, i
		
		' convert all pluses to spaces
		sOutput = replace(value, "+", " ",1,-1,1)
		
		' next convert %hexdigits to the character
		aSplit = Split(sOutput, "%")
		
		If IsArray(aSplit) Then
		  sOutput = aSplit(0)
		  For I = 0 to UBound(aSplit) - 1
			sOutput = sOutput & _
			  Chr("&H" & Left(aSplit(i + 1), 2)) &_
			  Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
		  Next
		End If
		
		URLDecode = sOutput
		
	End Function
		
		
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