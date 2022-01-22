<%

class cls_asplite_formbuilder

	private allFields, counter
	public postback,offSet,doScroll,id,onSubmit,reload,initialize,sameSession,target

	private sub class_initialize()

		set allFields		= aspl.dict
		onSubmit			= "aspAjax('POST',aspLiteAjaxHandler,$(this).serialize(),aspForm);return false;"
		counter				= 0
		offSet				= 150
		doScroll			= false 'true or false
		initialize			= true
		sameSession			= false

		'by default, a hidden field named "asplPostBack" is added to the collection of forms
		dim postBackHF : set postBackHF=field("hidden")
		postBackHF.add "name","asplPostBack"
		postBackHF.add "value","true"
		postBackHF.add "noinit","true"
		
		'postback=true in case the form has been submitted 
		if aspl.convertBool(aspL.getRequest("asplPostBack")) then
			postback=true
		else
			postback=false
		end if
		
		'by default, a hidden field named "asplSessionId" is added to the collection of fields
		'this can help to protect an application against CSRF (Cross Site Request Forgery)
		dim sameSessionHF : set sameSessionHF=field("hidden")
		sameSessionHF.add "name","asplSessionId"
		sameSessionHF.add "value",session.SessionID
		sameSessionHF.add "noinit","true"

		'check if a "known" sessionID is included in the returned form
		'sameSession can only be true AFTER a form was submitted!
		'if form.sameSession is true, you can be sure a given session is "known" by the server
		'this DOES NOT MEAN that session is SECURE!!! This has nothing to do with security.
		'But it can help to detect if a "postback" should be considered secure or not
		'Make sure to do additional checks on the session (eg: does it represent a logged-in user?)
		if postBack then
		
			if aspl.convertNmbr(aspL.getRequest("asplSessionId"))<>0 then
				if aspl.convertNmbr(aspL.getRequest("asplSessionId"))=aspl.convertNmbr(session.SessionID) then
					sameSession=true
				end if
			end if

		end if
		
		'if there is no asplTarget in the request-collections, aspLite assumes that the targetDIV=asplEvent
		'so the event that triggers an ASP function, then also defines wich id needs to be targetted
		if not aspl.isEmpty(aspl.getRequest("asplTarget")) then
			target=aspl.getRequest("asplTarget")
		else
			'if there is no asplTarget, aspLite assumes the ID of the calling DIV is the target
			target=aspl.getRequest("asplEvent")
		end if	
		
		'keep track of the targetDiv
		dim targetDIV : set targetDIV = field("hidden")
		targetDIV.add "name","asplTarget"
		targetDIV.add "value",target
		targetDIV.add "noinit","true"
		
		'automatically assign the ID using the targetDiv (unique ID)
		id=target & "_aspForm"		
		
		'this asplEvent-parameter is what triggers a specific ASP function to execute
		'this also has to be stored in the form so that a postback executes that same ASP function
		dim eventListener : set eventListener = field("hidden")
		eventListener.add "name","asplEvent"
		eventListener.add "value",aspl.getRequest("asplEvent")
		eventListener.add "noinit","true"		

	end sub	

	public function field(value)

		set field=aspl.dict
		field.add "type",value
		allFields.add counter,field
		counter=counter+1

	end function	

	public sub build()				
		
		'reload
		if aspl.convertNmbr(reload)>0 then

			if not postback then

				dim setTimeout: set setTimeout=field("script")
				setTimeout.add "text","$(document).ready(function(e) {setInterval(function(){$('#" & id & "').submit()}," & reload*1000 & ")})"

			end if

		end if	

		Dim arr : arr = Array()
		ReDim arr(counter-1)

		dim fieldkey
		
		'check if any given field is set to focus()
		for each fieldkey in allFields
			if allfields(fieldkey).exists("focus") then
				if not allfields(fieldkey).exists("id") then
					allfields(fieldkey).add "id",aspl.randomizer.randomText(10)
				end if
			end if			
		next		  
		for each fieldkey in allFields

			'set the values of the input fields with the request-values in the submitted form
			'only if is required to do so initialize=true
			
			if initialize then
				if not allfields(fieldkey).exists("noinit") then
					if allfields(fieldkey).exists("name") then
						if not aspl.isEmpty(aspL.getRequest(allFields(fieldkey)("name"))) then							
							allFields(fieldkey)("value")=aspL.getRequest(allFields(fieldkey)("name"))							
						end if
					end if
				end if
			end if

			set arr(fieldkey)=allFields(fieldkey)

		next		

		set allFields=nothing

		dim JsonAnswer,JsonHeader
		JsonAnswer=aspl.json.toJson("aspForm", arr, false)

		'finalizing JSON response - preparing header:
		JsonHeader = "{""target"":"""& target & ""","
		JsonHeader = JsonHeader & """offSet"":" & offSet & ","

		if doScroll then
			JsonHeader = JsonHeader & """doScroll"":true,"
		else
			JsonHeader = JsonHeader & """doScroll"":false,"
		end if

		JsonHeader = JsonHeader & """id"":""" & aspl.json.escape(id) & ""","
		JsonHeader = JsonHeader & """onSubmit"":""" & aspl.json.escape(onSubmit) & ""","
		JsonHeader = JsonHeader & """executionTime"":""[ASPLITE_executionTime]ms"","

		'removing from generated JSON initial bracket { and concatenating all together.
		JsonAnswer=right(JsonAnswer,Len(JsonAnswer)-1)
		JsonAnswer = JsonHeader & JsonAnswer

		'writing a response and stop executing page
		aspL.dumpJson JsonAnswer

	end sub

end class

%>