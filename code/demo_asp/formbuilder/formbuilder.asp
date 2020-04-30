<%
'This is a very basic ASP AJAX formbuilder application. The list of formfields is returned as a Json array, and turned
'into an html form by aspForm in /js/demoAjax.js. The built-in html5 input-types text, email, date, number and email are used.
'The form submission is also taken care of on this page (by reading out the postBack-hidden field)
'This is just an experiment. It needs a lot of development to make this a robust formbuilder. This is just a quick prototype.

class cls_formbuilder

	private allFields, counter
	public postback
	
	private sub class_initialize()
	
		set allFields=aspl.dict
		counter=0
		
		'by default, a hidden field named "postback" is added to the collection of forms
		dim postBackHF : set postBackHF=aspl.dict
		postBackHF.add "type","hidden"
		postBackHF.add "name","postBack"
		postBackHF.add "value","true"

		addField(postBackHF)	
		
		'set postback=true in case the forms has been submitted 
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
			if not aspl.isEmpty(aspL.getRequest(allFields(field)("name"))) then
				allFields(field)("value")=aspL.getRequest(allFields(field)("name"))
			end if
			
			set arr(field)=allFields(field)
			
		next
		
		set allFields=nothing

		json.dump (arr)	
	
	end sub

end class
%>