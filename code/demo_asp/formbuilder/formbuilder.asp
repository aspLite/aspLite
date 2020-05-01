<%
'This is a prototype, a sample application facilitated by aspLite

'The function addField expects a dictionary with one or more of the following key-item pairs:

' set dict=aspL.dict
' dict.add "type","XXX" - MANDATORY (possible values: text, email, date, number, submit, comment, select, radio, textarea)
' dict.add "name","firstname" - MANDATORY (except for comments)
' dict.add "html", "<strong>regular html</strong>", only used  for comments
' dict.add "label","First name:"
' dict.add "id","fname"
' dict.add "class","form-control"
' dict.add "maxlength",50 (numeric value)
' dict.add "required",true/false
' dict.add "style","color:Red"
' dict.add "value", the default value for a field
' dict.add "options",array of arrays!! (or a VBScript dictionary/ADO recordset in case of selectboxes)

'The list of formfields is returned as a Json array in "build()", and turned
'into an html form in /js/demoAjax.js. The built-in html5 input-types text, email, date, number and email are used.
'they are not supported in Safari... 


class cls_formbuilder

	private allFields, counter
	public postback,targetDiv,required,offSet
	
	private sub class_initialize()
	
		set allFields=aspl.dict
		counter=0
		required=" * required fields"
		offSet=100
		
		'by default, a hidden field named "postback" is added to the collection of forms
		dim postBackHF : set postBackHF=aspl.dict
		postBackHF.add "type","hidden"
		postBackHF.add "name","postBack"
		postBackHF.add "value","true"

		addField(postBackHF)	
		
		'postback=true in case the form has been submitted 
		if aspl.convertBool(aspL.getRequest("postBack")) then
			postback=true
		else
			postback=false
		end if
		
	end sub	
	
	public sub addField(dict)
	
		allFields.add counter,dict
		counter=counter+1
		
	end sub
	
	public sub build()
	
		Dim arr : arr = Array()
		ReDim arr(allFields.count-1)	

		for each field in allFields
			
			'set the values of the input fields with the request-values in the submitted form
			if allfields(field).exists("name") then
				if not aspl.isEmpty(aspL.getRequest(allFields(field)("name"))) then
					allFields(field)("value")=aspL.getRequest(allFields(field)("name"))
				end if
			end if
			
			set arr(field)=allFields(field)
			
		next
		
		set allFields=nothing

		JsonAnswer=json("aspForm", arr, false) 

		'finalizing JSON response - preparing header:
		JsonHeader = "{ ""targetDiv"": """& targetDiv &""", "
		JsonHeader = JsonHeader & """offSet"": " & offSet &", "& vbcrlf
		JsonHeader = JsonHeader & """required"": """ & json.escape(required) &""", "& vbcrlf
		  
		'removing from generated JSON initial bracket { and concatenating all together.
		JsonAnswer=right(JsonAnswer,Len(JsonAnswer)-1)
		JsonAnswer = JsonHeader & JsonAnswer		 		
		
		'writing a response and stop executing page
		aspL.dumpJson JsonAnswer	
		
	
	end sub

end class
%>