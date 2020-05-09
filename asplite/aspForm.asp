<%

class cls_asplite_formbuilder

	private allFields, counter, eventListener, p_target
	public postback,offSet,doScroll,id,onSubmit,reload,initialize

	private sub class_initialize()

		set allFields		= aspl.dict
		onSubmit			= "aspAjax('POST',aspLiteAjaxHandler,$(this).serialize(),aspForm);return false;"
		counter				= 0
		offSet				= 150
		doScroll			= false 'true or false
		set eventListener	= nothing
		initialize			= true

		'by default, a hidden field named "postback" is added to the collection of forms
		dim postBackHF : set postBackHF=field("hidden")
		postBackHF.add "name","postBack"
		postBackHF.add "value","true"

		'postback=true in case the form has been submitted 
		if aspl.convertBool(aspL.getRequest("postBack")) then
			postback=true
		else
			postback=false
		end if

	end sub

	public property get target

		target=p_target

	end property

	public property let target(value)

		p_target	= value

		'the form.id is automatically set, with a suffix "_aspForm"
		id			= p_target & "_aspForm"

	end property	

	public function field(value)

		set field=aspl.dict
		field.add "type",value
		allFields.add counter,field
		counter=counter+1

	end function

	public sub listenTo(eventName,eventValue)

		set eventListener = field("hidden")
		eventListener.add "name",eventName
		eventListener.add "value",eventValue

	end sub

	public sub build()
		
		'reload
		if aspl.convertNmbr(reload)>0 then

			if not postback then

				dim setTimeout: set setTimeout=field("script")
				setTimeout.add "text","$(document).ready(function(e) {setInterval(function(){$('#" & id & "').submit()}," & reload*1000 & ")})"

			end if

		end if
	

		Dim arr : arr = Array()
		ReDim arr(allFields.count-1)

		dim fieldkey
		for each fieldkey in allFields

			'set the values of the input fields with the request-values in the submitted form
			
			if initialize then
				if allfields(fieldkey).exists("name") then
					if not aspl.isEmpty(aspL.getRequest(allFields(fieldkey)("name"))) then
						allFields(fieldkey)("value")=aspL.getRequest(allFields(fieldkey)("name"))
					end if
				end if
			end if

			set arr(fieldkey)=allFields(fieldkey)

		next

		set allFields=nothing

		dim JsonAnswer,JsonHeader
		JsonAnswer=aspl.json.toJson("aspForm", arr, false)

		'finalizing JSON response - preparing header:
		JsonHeader = "{""target"":"""& p_target & ""","
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