<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit

dim asp
set asp=new cls_asp

class cls_asp	

	private startTime,stopTime,p_randomize,plugins	

	Private Sub Class_Initialize()
	
		startTime=Timer()
		
		p_randomize=false		
	
		Response.Buffer				= true
		session.Timeout				= 180 '3 hours
		server.ScriptTimeout		= 800 '3 minutes: needed for uploading bigger pictures/files!
		Response.CharSet			= "utf-8"
		Response.ContentType		= "text/html"
		Response.CacheControl		= "no-cache"
		Response.AddHeader "pragma", "no-cache"
		Response.Expires			= -1
		Response.ExpiresAbsolute	= Now()-1	

		set plugins=server.createobject("scripting.dictionary")
		
		
	End Sub	
	
	Private Sub Class_Terminate()
	
		'destroy all plugins
		dim p
		for each p in plugins
			set plugins(p)=nothing			
		next
	
		set plugins=nothing
		
	End sub
	

	public sub ASP_executeGlobal(path)

		path=lcase(path)
		
		dim strData
		strData=ASP_loadfile(path)
		
		strData=replace(strData,"<" & "%","",1,-1,1)
		strData=replace(strData,"%" & ">","",1,-1,1)	
		
		executeGlobal strData	

	end sub	
	
	public function ASP_loadfile(path)
		
		on error resume next

		Dim objStream
		Set objStream = server.CreateObject("ADODB.Stream")
		objStream.CharSet = "utf-8"
		objStream.Open	
		objStream.LoadFromFile(server.mappath(path))
		ASP_loadfile = objStream.ReadText()	
		set objStream=nothing
		
		if err.number<>0 then	
			asperror(path)
		end if	

	end function
	
	public function plugin(value)
	
		value=lcase(value)	
	
		if not plugins.exists(value) then
			
			ASP_executeGlobal("asp/plugins/" & value &".asp")	
			
			dim pluginCls
			set pluginCls=eval("new cls_asp_" & value)
			
			plugins.add value,pluginCls
		
		end if
		
		set plugin=plugins(value)
	
	end function
	
	private sub randomizer()
	
		'the function randomize() should only be called once/pageload
		
		if not p_randomize then 
			randomize()
			p_randomize=true
		end if	
	
	end sub
	
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
	
	public sub codebehind
		'this function is doing nothing, but leave it as it is
	end sub
	
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

	public function randomText(nmbrChars)	
	
		randomizer()

		dim i
		for i=1 to nmbrChars	
			randomText=randomText & CHR(Int((122-97+1)*Rnd+97))
		next
	   
	End Function

	public function randomNumber(startnr,stopnr)	
	
		randomizer()

		randomNumber=Int((stopnr-startnr+1)*Rnd+startnr)
		
	end function

	public function convert2 (byref getal)
		if len(getal)=1 then 
			convert2="0"&getal
		else
			convert2=getal
		end if
	end function
	
	Public Function printTimer() 	  
	   
		stopTime=Timer()
		
		PrintTimer = round((stopTime - startTime) * 1000,0) 'milliseconds
	
	End Function 
	

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
	
	private function asperror(value)
	
		response.clear
		
		asperror="<h1>Error  details:</h1>"
		asperror=asperror & value & "<br><br>"
		asperror=asperror & "err.number: " &  err.number & "<br><br>"
		asperror=asperror & "err.description: " &  err.description & "<br><br>"
				
		response.write asperror		
		
		response.end	
	
	end function
	
	public function responseServer (value)
	
		response.clear
		response.write value
		response.end	
	
	end function

end class

%>