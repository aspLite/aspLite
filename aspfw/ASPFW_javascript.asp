<%

class cls_javascript

	private headscripts, bodyscripts, script, counterHEAD,counterBODY
	
	Private Sub Class_Initialize()
	
		set headscripts=server.createobject("scripting.dictionary")
		set bodyscripts=server.createobject("scripting.dictionary")
		
		counterHEAD=0
		counterBODY=0
		
	End Sub
	
	Private Sub Class_Terminate()
	
		set headscripts=server.createobject("scripting.dictionary")
		set bodyscripts=server.createobject("scripting.dictionary")
		
	End Sub

	public function addHEAD(value)
		
		headscripts.add counterHEAD,value
		counterHEAD=counterHEAD+1		
	
	end function
	
	public function addBODY(value)
		
		bodyscripts.add counterBODY,value
		counterBODY=counterBODY+1		
	
	end function
	
	public function flushHEAD()
	
		flushHEAD="<script type=""text/javascript"">"		
	
		for each script in headscripts
		
			flushHEAD=flushHEAD & vbcrlf & vbtab & "// header script " & script+1 & "" & vbcrlf & vbtab & headscripts(script) & vbcrlf
		
		next
		
		flushHEAD=flushHEAD & "</script>"
		
		if flushHEAD="<script type=""text/javascript""></script>" then 
			flushHEAD=""
		else
			flushHEAD=vbcrlf & vbcrlf & "<!-- custom JS for this page -->" & vbcrlf & vbcrlf & flushHEAD & vbcrlf & vbcrlf & "<!-- end custom JS -->" & vbcrlf
		end if
	
	end function
	
	public function flushBODY()
	
		flushBODY="<script type=""text/javascript"">"		
	
		for each script in bodyscripts
		
			flushBODY=flushBODY & vbcrlf & vbtab & "// body script " & script+1 & "" & vbcrlf & vbtab & bodyscripts(script) & vbcrlf
		
		next
		
		flushBODY=flushBODY & "</script>"
		
		if flushBODY="<script type=""text/javascript""></script>" then 
			flushBODY=""
		else
			flushBODY=vbcrlf & vbcrlf & "<!-- custom JS for this page -->" & vbcrlf & vbcrlf & flushBODY & vbcrlf & vbcrlf & "<!-- end custom JS -->" & vbcrlf
		end if
	
	end function
	

end class
%>