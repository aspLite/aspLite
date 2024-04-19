class cls_asplite

	'VERSION: 1.0

	private debug,startTime,stopTime,plugins,p_fso,cacheprefix,multipart,p_json,p_randomizer,p_formmessages,p_value

	Private Sub Class_Initialize()

		on error resume next

		startTime					= Timer()
		debug						= aspL_debug

		'-------------------------------------------
		Response.Buffer				= true
		Response.CharSet			= "utf-8" 'does not work on IIS5 (Windows 2000), comment it out for IIS5
		Response.ContentType		= "text/html"
		Response.CacheControl		= "no-cache"
		Response.AddHeader "pragma", "no-cache"
		Response.Expires			= -1
		Response.ExpiresAbsolute	= Now()-1
		Server.ScriptTimeout		= 3600
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
		set p_formmessages	= nothing					  

		on error goto 0

	End Sub
	
	...