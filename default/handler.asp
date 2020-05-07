<%
dim html,body

'load the template of the demosite in the html-variable
'default.resx is plain html and it will remain 99% unchanged. 
'only [BODY] and [TIMER] will be replaced (see bottom of this file).
'however, in most cases, [BODY] will be empty, as aspLite will mainly
'return dynamic Ajax forms (see sampleform0 to 18)
html=aspL.loadText("default/html/default.resx")

'create a database object. It will also be used on various occasions in this demo
'this demo uses an Access database, located in "db/sample.mdb"
dim db : set db=aspL.plugin("database") : db.path="db/sample.mdb"

'here comes the eventhandler (what did the user do/click on exactly?)
'in general: if the code for dealing with the event is not large, include it here (downloadlargefile,downloadsmallfile,etc).
'if the code is getting complex, use yet another codebehind file, like all sampleforms!
'the aspl.exec is the DEFAULT sub in aspLite, so aspl(XX) will execute XX where XX is the path to the ASP file to be executed

select case lcase(aspL.getRequest("e")) 'e="event"	
	
	case "downloadlargefile" 		: aspL.dumpBinary("default/html/largefile.jpg")
	
	case "downloadsmallfile" 		: aspL.dumpBinary("default/html/smallfile.jpg")
				
	case "upload" 					: body=aspL.loadText("default/html/singleupload.resx")	
	
	case "uploadmulti" 				: body=aspL.loadText("default/html/multiupload.resx")	

	case "uploadfilejquery" 		: aspL("default/asp/uploadfile.asp") : aspL.die
	
	case "uploadfile" 				: aspL("default/asp/uploadfile.asp") ' takes care of the actual upload!
	
	case "body"						: aspL("default/asp/sampleform0.asp") 'default content for <div id="body">
		
	case "sampleform1" 				: aspL("default/asp/sampleform1.asp")
	
	case "sampleform2" 				: aspL("default/asp/sampleform2.asp")
	
	case "sampleform3"				: aspL("default/asp/sampleform3.asp")
	
	case "sampleform4" 				: aspL("default/asp/sampleform4.asp")	
	
	case "sampleform5" 				: aspL("default/asp/sampleform5.asp")	
	
	case "sampleform6" 				: aspL("default/asp/sampleform6.asp")	
	
	case "sampleform7" 				: aspL("default/asp/sampleform7.asp")
	
	case "sampleform8" 				: aspL("default/asp/sampleform8.asp")
	
	case "sampleform9" 				: aspL("default/asp/sampleform9.asp")
	
	case "sampleform10" 			: aspL("default/asp/sampleform10.asp")
	
	case "sampleform11" 			: aspL("default/asp/sampleform11.asp")

	case "sampleform12" 			: aspL("default/asp/sampleform12.asp")
	
	case "sampleform13" 			: aspL("default/asp/sampleform13.asp")	
	
	case "sampleform14" 			: aspL("default/asp/sampleform14.asp")	
	
	case "sampleform15" 			: aspL("default/asp/sampleform15.asp")	
	
	case "sampleform16" 			: aspL("default/asp/sampleform16.asp")	
	
	case "sampleform17" 			: aspL("default/asp/sampleform17.asp")
	
	case "sampleform18" 			: aspL("default/asp/sampleform18.asp")
	
	case "sampleform19" 			: aspL("default/asp/sampleform19.asp")	
	
	case "sampleform20" 			: aspL("default/asp/sampleform20.asp")
	
	case "sampleform20_data" 		: aspL("default/asp/datatables/sampleform20_data.asp")
	
	case "sampleform21" 			: aspL("default/asp/sampleform21.asp")	
	
	case "sampleform21_data" 		: aspL("default/asp/datatables/sampleform21_data.asp")	
	
	case "sampleform22" 			: aspL("default/asp/sampleform22.asp")
	
	case else 
				
		'get userfriendly url, if any (and launch a new handler-instance!)
		aspL("default/404.handler.asp")		
		
end select

'final replaces before returning html to the browser
'body
html=replace (html,"[BODY]",body,1,-1,1)

'timer
html=replace (html,"[TIMER]","<!--code took " & aspL.printTimer & " ms to execute-->",1,-1,1)

response.write html

'destroy aspLite
set aspL=nothing

%>