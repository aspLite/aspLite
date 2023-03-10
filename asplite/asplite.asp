<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
%>
<!-- #include file="config.asp"-->
<!-- #include file="aspForm.asp"-->
<%
dim aspL
set aspL=new cls_asplite

class cls_asplite

	'VERSION: 0.1.0

	private debug,startTime,stopTime,plugins,p_fso,cacheprefix,multipart,p_json,p_randomizer,p_value

	Private Sub Class_Initialize()

		on error resume next

		startTime					= Timer()
		debug						= aspL_debug

		'-------------------------------------------
		Response.Buffer				= true
		Response.CharSet			= "utf-8" 'does not work on IIS5 (Windows 2000 Servers) - comment it out when IIS5 is used!
		Response.ContentType		= "text/html"
		Response.CacheControl		= "no-cache"
		Response.AddHeader "pragma", "no-cache"
		Response.Expires			= -1
		Response.ExpiresAbsolute	= Now()-1
		Server.ScriptTimeout		= 600
		'-------------------------------------------

		'check if a form with enctype="multipart/form-data" was submitted. 
		'in that case, the request(.form) collection cannot be called as it throws an error
		'this is important for getRequest() -> see below
		If InStr(Request.ServerVariables("HTTP_CONTENT_TYPE"), "multipart")<>0 Then
			multipart=true
		else
			multipart=false
		end if

		cacheprefix="asplite_"
		set plugins			= nothing
		set p_fso			= nothing
		set p_json			= nothing
		set p_randomizer	= nothing

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

	'default sub!
	'executeGlobal: the ASP code in the file (path) will
	'be executed as classic ASP and it will be available in
	'the namespace of this page.
	public default sub exec(path)

		on error resume next

		executeGlobal removeCRB(stream(path,false,""))
		
		aspError("problem when executing " & path)

		on error goto 0

	end sub
	
	public function recycleApplication

		on error resume next
		
		err.clear
		
		'this function tries to update the global.asa file in the root of your site (if any). 
		'this automatically recycles the IIS application pool

		dim globalAsaPath : globalAsaPath=server.mappath("/global.asa")
		dim globalAsaText : globalAsaText=fso.opentextfile(globalAsaPath).readAll
		dim globalAsaNew : set globalAsaNew=fso.CreateTextFile(globalAsaPath)
		globalAsaNew.write(globalAsaText) : globalAsaNew.close : set globalAsaNew=nothing
		
		if err.number<>0 then
			recycleApplication=false
		else
			recycleApplication=true
		end if
		
		err.clear
		
		on error goto 0

	end function
	
	public Function SaveBinaryData(FileName, ByteArray)
	
		'Create Stream object
		Dim BinaryStream : Set BinaryStream = server.createObject("ADODB.Stream")

		'Specify stream type - we want To save binary data.
		BinaryStream.Type = 1 'adTypeBinary

		'Open the stream And write binary data To the object
		BinaryStream.Open
		BinaryStream.Write ByteArray

		'Save binary data To disk
		BinaryStream.SaveToFile FileName, 2 'adSaveCreateOverWrite
		
	End Function
	
	
	public Function saveAsFile(fileName, fileBody)
	
		response.clear	
		Response.AddHeader "Content-Disposition", "attachment; filename=" & safeFileName(filename)
		response.write fileBody		
		response.flush()
		response.clear

		die()
		
	end function
	
	public function safeFileName(value)
	
		safeFileName=server.urlEncode(convertStr(value))
		safeFileName=replace(safeFileName,"+"," ",1,-1,1)	
	
	end function

	public function removeCRB(value)

		'remove Code Render Blocks

		value=replace(value,"<" & "%","",1,-1,1)
		removeCRB=replace(value,"%" & ">","",1,-1,1)

	end function

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

	public function form

		set form=new cls_asplite_formbuilder

	end function

	public function json

		'gives back a json object (created only once)

		if p_json is nothing then
			set p_json=new cls_asplite_json
		end if

		set json=p_json

	end function

	public function randomizer

		'gives back a randomizer object (created only once)

		if p_randomizer is nothing then
			set p_randomizer=new cls_asplite_randomizer
		end if

		set randomizer=p_randomizer

	end function

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

		err.clear()

		if multipart then

			'binary data was submitted with enctype=multipart/form-data
			'in this case, the request.form collecion cannot be used and even raises an error
			'aspLite only returns the querystring in this case

			getRequest=request.querystring(value)

		else

			if not [isEmpty](request.form(value)) then
				getRequest=request.form(value)
			elseif [isEmpty](request.querystring(value)) then
				getRequest=request.querystring(value)
			else
				getRequest=request(value)
			end if

		End If

		on error goto 0

	end function

	Public function asperror(value)

		if not debug then exit function

		if err.number<>0 then

			asperror="aspLite error details:" & vbcrlf			
			asperror=asperror & "url: " & curPageURL & vbcrlf 
			asperror=asperror & "querystring: " & Request.ServerVariables("QUERY_STRING") & vbcrlf 
			asperror=asperror & "form: " & Request.form & vbcrlf 
			asperror=asperror & "error: " & value & vbcrlf 
			asperror=asperror & "err.number: " & err.number & vbcrlf 
			asperror=asperror & "err.source: " & err.source & vbcrlf 
			asperror=asperror & "err.description: " & err.description

			json.dump(asperror)

		end if

	end function

	public sub die

		Class_Terminate()
		response.end

	end sub

	public function dump (value)

		response.clear
		response.write replace(value,"[ASPLITE_executionTime]",printTimer,1,-1,1)
		die()

	end function

	public function dumpJson (value)

		Response.ContentType = "application/json"
		dump(value)

	end function

	public function dumpBinary (path,dumpAs)

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
		'retrieve filename
		dim filename
		if not [isEmpty](dumpAs) then
			filename=dumpAs
		else
			filename=right(path,len(path)-InStrRev(path,"\",-1,1))
		end if			  

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
			case "rtf" : response.ContentType="text/rtf"
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

		Response.AddHeader "Content-Disposition", "attachment; filename=" & safeFileName(filename)

		if size<chunksize then
			response.AddHeader "Content-Length", size
			response.binarywrite objStream.Read()
			response.flush()	   
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
	'### some logging functions
	'############################	
	
	public function log(value)
	
		dim logfile 
		logfile=loadText(aspL_path & "/asplite.log")
	
		saveText aspL_path & "/asplite.log",logfile & now & vbcrlf & value &  vbcrlf & "FORM elements: " & vbcrlf & formcollection & vbcrlf
	
	end function
	
	private function formcollection
	
		exit function
		
		dim element	
		for each element in request.form
			formcollection=formcollection & element & ": " & request.form(element) & vbcrlf
		next
		
		for each element in request.querystring
			formcollection=formcollection & element & ": " & request.querystring(element) & vbcrlf
		next
	
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

		if not [isEmpty](getCacheT) then
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

		if not [isEmpty](ufl) then

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

			if not [isEmpty](ufl) and IsAlphaNumeric(ufl) then
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

	'converts a VBScript date to YYYY-MM-DD
	'this is one of these things that should have been added to VBScript after the 2010-adoption of HTML5
	'how difficult can it be to add this to the list of default FormatDateTime options ???!!!
	public function convertHtmlDate(value)
	
		if [isEmpty](value) then
			convertHtmlDate=""
		end if
	
		if isDate(value) then
			convertHtmlDate=year(value) & "-" & padleft(month(value),2,0) & "-" & padleft(day(value),2,0)
		else
			convertHtmlDate=""
		end if
	
	
	end function
	
	'get file-extension of a file (eg: jpeg, jpg, gif, doc, docx, etc)
	public function getFileType(filename)

		getFileType=right(filename,len(filename)-InStrRev(filename,".",-1,1))

	end function

	public function isNumber(byval value)

		on error resume next

		if isNull(value) then 
			isNumber=false
		elseif [isEmpty](value) then
			isNumber=false
		else
			isNumber=isNumeric(value)
		end if

		on error goto 0

	end function
	
	public function length(value)
	
		on error resume next

		if [isEmpty](value) then
			length=0
		else
			length=len(value)
		end if

		on error goto 0
	
	end function	
	
	public function [isEmpty](byval value)

		on error resume next

		isEmpty=false

		if isNull(value) then
			isEmpty=true
		else
			if trim(value)="" then isEmpty=true
		end if

		on error goto 0

	End Function

	public function convertStr(value)

		p_value=value

		if not isnull(p_value) then
			convertStr=trim(cstr(p_value))
		else
			convertStr=""
		end if

	End Function
	
	public function convertNull(value)

		p_value=value		
		
		if isNull(p_value) then
		
			convertNull=null
			exit function
			
		elseif convertNmbr(p_value)=0 then
			
			convertNull=null
			exit function
		
		elseif isNumber(p_value) then
		
			convertNull=convertNmbr(p_value)
			exit function
			
		elseif [isEmpty](p_value) then
		
			convertNull=null
			exit function
			
		else
		
			convertNull=p_value
			
		end if

	End Function

	public function sanitize(sValue)

		if [isEmpty](sValue) then
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
	
	public function htmlEncJs (sValue)
	
		htmlEncJs=sanitizeJS(server.htmlEncode(sValue))
	
	end function	 

	public function convertNmbr(value)
	
		dim cValue : cValue=value
		
		on error resume next

		if isNumber(cValue) then 
			convertNmbr=cdbl(cValue)
		else
			convertNmbr=0
		end if

		if err.number<>0 then convertNmbr=0

		on error goto 0

	End Function
	

	public function convertBool(value)

		On Error Resume Next

		if [isEmpty](value) then
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
		if [isEmpty](str) then
			sqli=""
		else
			sqli=replace(str,"'","''",1,-1,1)
		end if
	end function
	

	'******************************************************************************************
	'* padLeft - copied from Ajaxed Library
	'******************************************************************************************
	public function padLeft(value, totalLength, paddingChar)
		padLeft = right(clone(paddingChar, totalLength) & value, totalLength)
	end function

	'******************************************************************************************
	'* clone - copied from Ajaxed Library
	'******************************************************************************************
	public function clone(byVal str, n)
		dim i
		for i = 1 to n : clone = clone & str : next
	end function

	public Function IsAlphaNumeric(byVal str)

		If IsNull(str) Then str = ""

		Dim ianRegEx
		Set ianRegEx = New RegExp
		ianRegEx.Pattern = "[^a-z0-9\/\_\-\.]"
		ianRegEx.Global = True
		ianRegEx.IgnoreCase = True
		IsAlphaNumeric = (str = ianRegEx.Replace(str,""))

		Set ianRegEx=nothing

	End Function

	'proper case -> first character in names will convert to uppercase, others to lowercase
	'this wil also work for double names like Scott-Johnson or O'Connor.
	public function pCase(value)
	
		pcase=""
	
		value=aspl.convertStr(value)
		if [isEmpty](value) then exit function
		
		Dim i, x, strOut, flg
		flg = True
		For i = 1 To Len(value)
		  x = LCase(Mid(value, i, 1))
		  If Not IsNumeric(x) And (x < "a" Or x > "z") and (x < "à" or x > "þ")  Then
			flg = True
		  ElseIf flg Then
			x = UCase(x)
			flg = False
		  End If
		  strOut = strOut & x
		Next
		
		PCase = strOut

	end function
	
	
	public function curPageURL()
	
		dim s, protocol, port

		if Request.ServerVariables("HTTPS") = "on" then
			s = "s"
		else
			s = ""
		end if  

		protocol = strleft(LCase(Request.ServerVariables("SERVER_PROTOCOL")), "/") & s 

		if Request.ServerVariables("SERVER_PORT") = "80" then
			port = ""
		else
			port = ":" & Request.ServerVariables("SERVER_PORT")
		end if  

		curPageURL = protocol & "://" & Request.ServerVariables("SERVER_NAME") &_
		port & Request.ServerVariables("SCRIPT_NAME")
		
	end function
	

	public function strLeft(str1,str2)
		strLeft = Left(str1,InStr(str1,str2)-1)
	end function
	
	
	public Function GetEmailValidator()

		Set GetEmailValidator = New RegExp

		GetEmailValidator.Pattern = "^((?:[A-Z0-9_%+-]+\.?)+)@((?:[A-Z0-9-]+\.)+[A-Z]{2,40})$"

		GetEmailValidator.IgnoreCase = True

	End Function

	' Action: checks if an email is correct. 
	' Parameter: Email address 
	' Returned value: on success it returns True, else False. 
	public Function CheckEmailSyntax(strEmail) 
	
		Dim EmailValidator : Set EmailValidator = GetEmailValidator()

		CheckEmailSyntax=EmailValidator.Test(strEmail)
	 
	end function
	
	
	public function stripHTML(strHTML)
		'Strips the HTML tags from strHTML

		Dim objRegExp, strOutput
		Set objRegExp = New Regexp

		objRegExp.IgnoreCase = True
		objRegExp.Global = True
		objRegExp.Pattern = "<(.|\n)+?>"

		'Replace all HTML tag matches with the empty string
		strOutput = objRegExp.Replace(strHTML, "")

		'Replace all < and > with &lt; and &gt;
		strOutput = Replace(strOutput, "<", "&lt;")
		strOutput = Replace(strOutput, ">", "&gt;")

		stripHTML = strOutput    'Return the value of strOutput

		Set objRegExp = Nothing
	End function

end class

'**************************************************************************************************************

'' @CLASSTITLE:		JSON
'' @CREATOR:		Michal Gabrukiewicz (gabru at grafix.at), Michael Rebec
'' @CONTRIBUTORS:	- Cliff Pruitt (opensource at crayoncowboy.com)
''					- Sylvain Lafontaine
''					- Jef Housein
''					- Jeremy Brown
'' @CREATEDON:		2007-04-26 12:46
'' @CDESCRIPTION:	Comes up with functionality for JSON (http://json.org) to use within ASP.
''					Correct escaping of characters, generating JSON Grammer out of ASP datatypes and structures
''					Some examples (all use the <em>toJSON()</em> method but as it is the class' default method it can be left out):
''					<code>
''					<%
''					'simple number
''					output = (new JSON)("myNum", 2, false)
''					'generates {"myNum": 2}
''
''					'array with different datatypes
''					output = (new JSON)("anArray", array(2, "x", null), true)
''					'generates "anArray": [2, "x", null]
''					'(note: the last parameter was true, thus no surrounding brackets in the result)
''					% >
''					</code>
'' @REQUIRES:		-
'' @OPTIONEXPLICIT:	yes
'' @VERSION:		1.5.1

'**************************************************************************************************************
class cls_asplite_json

	'private members
	private output, innerCall, bufferCount

	'public members
	public toResponse		''[bool] should the generated representation be written directly to the response (using <em>response.write</em>)? default = false
	public recordsetPaging	''[bool] indicates if only the current page should be processed on paged recordsets.
							''e.g. would return only 10 records if <em>RS.pagesize</em> is set to 10. default = false (means that always all records are processed)

	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		newGeneration()
		toResponse = false
		recordsetPaging = false
		bufferCount=0	   
	end sub

	'******************************************************************************************
	'' @SDESCRIPTION:	STATIC! takes a given string and makes it JSON valid
	'' @DESCRIPTION:	all characters which needs to be escaped are beeing replaced by their
	''					unicode representation according to the 
	''					RFC4627#2.5 - http://www.ietf.org/rfc/rfc4627.txt?number=4627
	'' @PARAM:			val [string]: value which should be escaped
	'' @RETURN:			[string] JSON valid string
	'******************************************************************************************
	public function escape(val)
		dim cDoubleQuote, cRevSolidus, cSolidus
		cDoubleQuote = &h22
		cRevSolidus = &h5C
		cSolidus = &h2F
		dim i, currentDigit
		for i = 1 to (len(val))
			currentDigit = mid(val, i, 1)
			if ascw(currentDigit) > &h00 and ascw(currentDigit) < &h1F then
				currentDigit = escapequence(currentDigit)
			elseif ascw(currentDigit) >= &hC280 and ascw(currentDigit) <= &hC2BF then
				currentDigit = "\u00" + right(padLeft(hex(ascw(currentDigit) - &hC200), 2, 0), 2)
			elseif ascw(currentDigit) >= &hC380 and ascw(currentDigit) <= &hC3BF then
				currentDigit = "\u00" + right(padLeft(hex(ascw(currentDigit) - &hC2C0), 2, 0), 2)
			else
				select case ascw(currentDigit)
					case cDoubleQuote: currentDigit = escapequence(currentDigit)
					case cRevSolidus: currentDigit = escapequence(currentDigit)
					case cSolidus: currentDigit = escapequence(currentDigit)
				end select
			end if
			escape = escape & currentDigit
		next
	end function

	'******************************************************************************************************************
	'' @SDESCRIPTION:	generates a representation of a name value pair in JSON grammer
	'' @DESCRIPTION:	It generates a name value pair which is represented as <em>{"name": value}</em> in JSON.
	''					the generation is fully recursive. Thus the value can also be a complex datatype (array in dictionary, etc.) e.g.
	''					<code>
	''					<%
	''					set j = new JSON
	''					j.toJSON "n", array(RS, dict, false), false
	''					j.toJSON "n", array(array(), 2, true), false
	''					% >
	''					</code>
	'' @PARAM:			name [string]: name of the value (accessible with javascript afterwards). leave empty to get just the value
	'' @PARAM:			val [variant], [int], [float], [array], [object], [dictionary], [recordset]: value which needs
	''					to be generated. Conversation of the data types is as follows:<br>
	''					- <strong>ASP datatype -> JavaScript datatype</strong>
	''					- NOTHING, NULL -> null
	''					- INT, DOUBLE -> number
	''					- STRING -> string
	''					- BOOLEAN -> bool
	''					- ARRAY -> array
	''					- DICTIONARY -> Represents it as name value pairs. Each key is accessible as property afterwards. json will look like <code>"name": {"key1": "some value", "key2": "other value"}</code>
	''					- <em>multidimensional array</em> -> Generates a 1-dimensional array (flat) with all values of the multidimensional array
	''					- RECORDSET -> array where each row of the recordset represents a field of the array. Each array field has properties according to the column names of the recordset (<strong>lowercase!</strong>) e.g <em>toJSON("r", RS, false)</em> can be accessed afterwards with <em>r[0].id</em>
	''					- <em>request</em> object -> every property and collection (cookies, form, querystring, etc) of the asp request object is exposed as an item of a dictionary. Property names are <strong>lowercase</strong>. e.g. <em>servervariables</em>.
	''					- OBJECT -> name of the type (if unknown type) or all its properties (if class implements <em>reflect()</em> method)
	''					Implement a <strong>reflect()</strong> function if you want your custom classes to be recognized. The function must return
	''					a dictionary where the key holds the property name and the value its value. Example of a reflect function within a User class which has firstname and lastname properties
	''					<code>
	''					<%
	''					function reflect()
	''					.	set reflect = server.createObject("scripting.dictionary")
	''					.	reflect.add "firstname", firstname
	''					.	reflect.add "lastname", lastname
	''					end function
	''					% >
	''					</code>
	''					Example of how to generate a JSON representation of the asp request object and access the <em>HTTP_HOST</em> server variable in JavaScript:
	''					<code>
	''					<script>alert(<%= (new JSON)(empty, request, false) % >.servervariables.HTTP_HOST);</script>
	''					</code>
	'' @PARAM:			nested [bool]: indicates if the name value pair is already nested within another? if yes then the <em>{}</em> are left out.
	'' @RETURN:			[string] returns a JSON representation of the given name value pair
	''					(if toResponse is on then the return is written directly to the response and nothing is returned)
	'******************************************************************************************************************
	public default function toJSON(name, val, nested)

		if not nested and not isEmpty(name) then write("{")
		if not isEmpty(name) then write("""" & escape(name) & """: ")
		generateValue(val)
		if not nested and not isEmpty(name) then write("}")
		toJSON = output

		if innerCall = 0 then newGeneration()
	end function

	'******************************************************************************************************************
	'' Added by Pieter Cooreman, developer aspLite
	'' I added this function to have a very easy way to return single level values named "aspl"
	'' Val can be of all types (string, arrays, objects, recordset) - see toJSON
	'' Another reason for this "hack" is that toResponse=true ensures that the variable "output" does not 
	'' get too big (it' not even used). Adding strings to ever growing string-variables results in very poor
	'' performance very quickly! Having the response.buffer=true while adding many strings to it,
	'' gives much better performance than classic string-concatenation
	'******************************************************************************************************************
	public sub dump(val)
	
		toResponse=true
	
		Response.ContentType = "application/json"
		
		response.clear

		write("{""aspl"": ")
		generateValue(val)
		write("}")
		
		response.end

	end sub

	'******************************************************************************************************************
	'* generate 
	'******************************************************************************************************************
	private function generateValue(val)
		if isNull(val) then
			write("null")
		elseif isArray(val) then
			generateArray(val)
		elseif isObject(val) then
			dim tName : tName = typename(val)
			if val is nothing then
				write("null")
			elseif tName = "Dictionary" or tName = "IRequestDictionary" then
				generateDictionary(val)
			elseif tName = "Recordset" then
				generateRecordset(val)
			elseif tName = "IRequest" then
				set req = aspL.dict
				req.add "clientcertificate", val.ClientCertificate
				req.add "cookies", val.cookies
				req.add "form", val.form
				req.add "querystring", val.queryString
				req.add "servervariables", val.serverVariables
				req.add "totalbytes", val.totalBytes
				generateDictionary(req)
			elseif tName = "IStringList" then
				if val.count = 1 then
					toJSON empty, val(1), true
				else
					generateArray(val)
				end if
			else
				generateObject(val)
			end if
		else
			'bool
			dim varTyp
			varTyp = varType(val)

			select case aspL.convertNmbr(varTyp)

				case 11 : if val then write("true") else write("false") 'bool

				case 2,3,17,19 : write(cLng(val))

				case 4,5,6,14 : write(replace(cDbl(val), ",", "."))

				case else :	write("""" & escape(val & "") & """")

			end select 

		end if
		generateValue = output
	end function

	'******************************************************************************************************************
	'* generateArray 
	'******************************************************************************************************************
	private sub generateArray(val)
		dim item, i
		write("[")
		i = 0
		'the for each allows us to support also multi dimensional arrays
		for each item in val
			if i > 0 then write(",")
			generateValue(item)
			i = i + 1
		next
		write("]")
	end sub

	'******************************************************************************************************************
	'* generateDictionary 
	'******************************************************************************************************************
	private sub generateDictionary(val)
		innerCall = innerCall + 1
		if val.count = 0 then
			toJSON empty, null, true
			exit sub
		end if
		dim key, i
		write("{")
		i = 0
		for each key in val
			if i > 0 then write(",")
			toJSON key, val(key), true
			i = i + 1
		next
		write("}")
		innerCall = innerCall - 1
	end sub

	'******************************************************************************************************************
	'* generateRecordset 
	'******************************************************************************************************************
	private sub generateRecordset(val)
		dim i, curRow,copyDate
		write("[")
		curRow = 0
		'recordset.pagesize = -1 means it is not paged.
		while not val.eof and ((recordsetPaging and curRow < val.pageSize) or val.recordCount = -1 or not recordsetPaging)
			innerCall = innerCall + 1
			write("{")
			for i = 0 to val.fields.count - 1
				if i > 0 then write(",")

				if isDate(val.fields(i).value) then
					copyDate=val.fields(i).value
					toJSON val.fields(i).name, year(copyDate) & "-" & padLeft(month(copyDate),2,0) & "-" & padLeft(day(copyDate),2,0) & "T" & padLeft(hour(copyDate),2,0) & ":" & padLeft(minute(copyDate),2,0) & ":" & padLeft(second(copyDate),2,0), true
				else
					toJSON val.fields(i).name, val.fields(i).value, true
				end if
			next
			write("}")
			val.movenext()
			curRow = curRow + 1
			if not val.eof and ((recordsetPaging and curRow < val.pageSize) or val.recordCount = -1 or not recordsetPaging) then write(",")
			innerCall = innerCall - 1
		wend
		write("]")
	end sub

	'******************************************************************************************************************
	'* generateObject 
	'******************************************************************************************************************
	private sub generateObject(val)
		dim props
		on error resume next
		set props = val.reflect()
		if err = 0 then
			on error goto 0
			innerCall = innerCall + 1
			toJSON empty, props, true
			innerCall = innerCall - 1
		else
			on error goto 0
			write("""" & escape(typename(val)) & """")
		end if
	end sub

	'******************************************************************************************************************
	'* newGeneration 
	'******************************************************************************************************************
	private sub newGeneration()
		output = empty
		innerCall = 0
	end sub

	'******************************************************************************************
	'* JsonEscapeSquence 
	'******************************************************************************************
	private function escapequence(digit)
		escapequence = "\u00" + right(padLeft(hex(ascw(digit)), 2, 0), 2)
	end function

	'******************************************************************************************
	'* padLeft 
	'******************************************************************************************
	private function padLeft(value, totalLength, paddingChar)
		padLeft = right(clone(paddingChar, totalLength) & value, totalLength)
	end function

	'******************************************************************************************
	'* clone 
	'******************************************************************************************
	private function clone(byVal str, n)
		dim i
		for i = 1 to n : clone = clone & str : next
	end function

	'******************************************************************************************
	'* write 
	'******************************************************************************************
	private sub write(val)
		if toResponse then
			bufferCount=bufferCount+1				
			response.write(val)
			if bufferCount>1000 then 
				'emptying the buffer after some response.writes improves performance
				response.flush				
				bufferCount=0
			end if															
		else
			output = output & val
		end if
	end sub

end class

class cls_asplite_randomizer

	Private Sub Class_Initialize()

		randomize()

	end sub

	public function randomText(nmbrChars)

		dim i
		for i=1 to nmbrChars
			randomText=randomText & CHR(Int((26)*Rnd+97))
		next

	End Function

	public function randomNumber(startnr,stopnr)

		randomNumber=aspl.convertNmbr((stopnr-startnr+1)*Rnd+startnr)

	end function
	
	public function CreateGUID(tmpLength)
		Randomize Timer
		Dim tmpCounter,tmpGUID,strValid
		strValid = "0123456789abcdefghijklmnopqrstuvwxyz"
		For tmpCounter = 1 To tmpLength
			tmpGUID = tmpGUID & Mid(strValid, Int(Rnd(1) * Len(strValid)) + 1, 1)
		Next
		CreateGUID = tmpGUID
	end function

end class
%>
