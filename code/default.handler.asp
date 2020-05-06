<%
dim html,body

'load the template of the demosite
html=aspL.loadText("code/html/default.resx")

'create a json object for all AJAX communications with the client
'be careful, this object will be used all over this demo site! do not destroy it!
dim json : set json=aspL.json

'create a database object. It will also be used on various occasions in this demo
dim db : set db=aspL.plugin("database") : db.path="db/sample.mdb"

'here you can typically add some sort of eventhandler (what did the user do exactly?)
'in general: if the code for dealing with the event is not large, include it here.
'if the code is getting complex, use yet another codebehind file, like in the examples below (helloworld, mail, json, etc)

select case lcase(aspL.getRequest("e")) '"event"	
	
	case "downloadlargefile" 		: aspL.dumpBinary("code/html/largefile.jpg")
	
	case "downloadsmallfile" 		: aspL.dumpBinary("code/html/smallfile.jpg")
	
	case "datatables_ssp"			: body=aspL.loadText("code/html/datatables_ssp.resx")
	
	case "datatables" 				: body=aspL.loadText("code/html/datatables.resx")

	case "json_datatables_data" 	: aspL.exec("code/asp/datatables_ssp.asp")	
	
	case "json_datatables" 			: aspL.exec("code/asp/json_datatables.asp")	
				
	case "upload" 					: body=aspL.loadText("code/html/singleupload.resx")
		
	case "uploadfile" 				: aspL.exec("code/asp/uploadfile.asp")	
	
	case "uploadmulti" 				: body=aspL.loadText("code/html/multiupload.resx")		
	
	'AJAX handlers
	'json.dump is very often used for AJAX handlers. "dump" basically means "flush" and "stop"	
	'All communication with the browser is done with JSON.
	
	case "sampleform0"				: aspL.exec("code/asp/sampleform0.asp")
		
	case "sampleform1" 				: aspL.exec("code/asp/sampleform1.asp")
	
	case "sampleform2" 				: aspL.exec("code/asp/sampleform2.asp")
	
	case "sampleform3"				: aspL.exec("code/asp/sampleform3.asp")
	
	case "sampleform4" 				: aspL.exec("code/asp/sampleform4.asp")	
	
	case "sampleform5" 				: aspL.exec("code/asp/sampleform5.asp")
	
	case "sampleform6" 				: aspL.exec("code/asp/sampleform6.asp")	
	
	case "sampleform7" 				: aspL.exec("code/asp/sampleform7.asp")
	
	case "sampleform8" 				: aspL.exec("code/asp/sampleform8.asp")
	
	case "sampleform9" 				: aspL.exec("code/asp/sampleform9.asp")
	
	case "sampleform10" 			: aspL.exec("code/asp/sampleform10.asp")
	
	case "sampleform11" 			: aspL.exec("code/asp/sampleform11.asp")

	case "sampleform12" 			: aspL.exec("code/asp/sampleform12.asp")
	
	case "sampleform13" 			: aspL.exec("code/asp/sampleform13.asp")	
	
	case "sampleform14" 			: aspL.exec("code/asp/sampleform14.asp")	
	
	case "sampleform15" 			: aspL.exec("code/asp/sampleform15.asp")	
	
	case "sampleform16" 			: aspL.exec("code/asp/sampleform16.asp")	
	
	case "sampleform17" 			: aspL.exec("code/asp/sampleform17.asp")	
	
	case "postdtt" 					: aspL.exec("code/asp/postdtt.asp")	
	
	case "uploadfilejquery" 		: aspL.exec("code/asp/uploadfile.asp") : aspL.die	''uploader
	
	case else 
				
		'get userfriendly url, if any (and launch a new handler-instance!)
		aspL.exec("code/404.handler.asp")		
		
end select

%>