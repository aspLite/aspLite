<%
'This is a prototype, a sample application facilitated by aspLite

' set dict=aspL.dict
' dict.add "type","XXX" - MANDATORY (possible values: text, email, date, number, submit, element, select, radio, textarea, script, button)
' dict.add "name","firstname"
' dict.add "html", "<strong>regular html</strong>", only used  for comments
' dict.add "label","First name:"
' dict.add "id","fname"
' dict.add "class","form-control"
' dict.add "maxlength",50 (numeric value)
' dict.add "required",true/false
' dict.add "style","color:Red"
' dict.add "value", the default value for a field
' dict.add "options", VBScript dictionary

'The list of formfields is returned as a Json array in "build()", and turned
'into an html form in /js/demoAjax.js. The built-in html5 input-types text, email, date, number and email are used.
'they are not supported in Safari... 

class cls_formbuilder

	private allFields, counter, eventListener
	public postback,targetDiv,offSet,doScroll,id,onSubmit
	
	private sub class_initialize()
	
		set allFields		= aspl.dict
		onSubmit			= "aspAjax('POST','',$(this).serialize(),aspForm);return false;"		
		counter				= 0	
		offSet				= 150
		doScroll			= false 'true or false	
		set eventListener = nothing
		
		'by default, a hidden field named "postback" is added to the collection of forms
		dim postBackHF : set postBackHF=field
		postBackHF.add "type","hidden"
		postBackHF.add "name","postBack"
		postBackHF.add "value","true"	
		
		'postback=true in case the form has been submitted 
		if aspl.convertBool(aspL.getRequest("postBack")) then
			postback=true
		else
			postback=false
		end if
		
	end sub	
	
	public function field
	
		set field=aspl.dict
		allFields.add counter,field
		counter=counter+1
		
	end function
	
	public sub listenTo(eventName,eventValue)	

		set eventListener	= aspl.dict
		eventListener.add "type","hidden"		
		eventListener.add "name",eventName
		eventListener.add "value",eventValue		
	
	end sub
	
	public sub build()
	
		Dim arr : arr = Array()
		ReDim arr(allFields.count)	

		for each fieldkey in allFields
			
			'set the values of the input fields with the request-values in the submitted form
			if allfields(fieldkey).exists("name") then
				if not aspl.isEmpty(aspL.getRequest(allFields(fieldkey)("name"))) then
					allFields(fieldkey)("value")=aspL.getRequest(allFields(fieldkey)("name"))
				end if
			end if
			
			set arr(fieldkey)=allFields(fieldkey)
			
		next
		
		'pass the event listener (if any) - this will always be the last item in the dictionary of fields
		if not eventListener is nothing then 
			set arr(allFields.count)=eventListener
		else
			ReDim preserve arr(allFields.count-1)	
		end if
		
		set allFields=nothing

		JsonAnswer=json("aspForm", arr, false) 

		'finalizing JSON response - preparing header:
		JsonHeader = "{""targetDiv"":"""& targetDiv & ""","
		JsonHeader = JsonHeader & """offSet"":" & offSet & ","
		
		if doScroll then
			JsonHeader = JsonHeader & """doScroll"":true,"
		else
			JsonHeader = JsonHeader & """doScroll"":false,"
		end if		
		
		JsonHeader = JsonHeader & """id"":""" & json.escape(id) & ""","
		JsonHeader = JsonHeader & """onSubmit"":""" & json.escape(onSubmit) & ""","
				  
		'removing from generated JSON initial bracket { and concatenating all together.
		JsonAnswer=right(JsonAnswer,Len(JsonAnswer)-1)
		JsonAnswer = JsonHeader & JsonAnswer		 		
		
		'writing a response and stop executing page
		aspL.dumpJson JsonAnswer		
	
	end sub

end class
%>