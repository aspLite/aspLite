<%
dim html,body

'load the template of the demosite
html=aspL.loadText("html/demo_asp/demo.resx")

'create a json object for all AJAX communications with the client
'be careful, this object will be used all over this demo site! do not destroy it!
dim json : set json=aspL.plugin("json")

'create a database object. It will also be used in the demo site all over the place.
dim db : set db=aspL.plugin("database") : db.path="db/sample.mdb"

'we will use the formbuilder all over the place, so we load the formbuilder class in the namespace
'this is very similar to an include directive.
aspl.exec("code/demo_asp/formbuilder/formbuilder.asp")

'here you can typically add some sort of eventhandler (what did the user do exactly?)
'in general: if the code for dealing with the event is not large, include it here.
'if the code is getting complex, use yet another codebehind file, like in the examples below (helloworld, mail, json, etc)

select case lcase(aspL.getRequest("e")) '"event"	
	
	case "downloadlargefile" 		: aspL.dumpBinary("html/demo_asp/largefile.jpg")
	
	case "downloadsmallfile" 		: aspL.dumpBinary("html/demo_asp/smallfile.jpg")
	
	case "datatables_ssp"			: body=aspL.loadText("html/demo_asp/datatables_ssp.resx")
	
	case "datatables" 				: body=aspL.loadText("html/demo_asp/datatables.resx")

	case "json_datatables_data" 	: aspL.exec("code/demo_asp/datatables_ssp.asp")	
	
	case "json_datatables" 			: aspL.exec("code/demo_asp/json_datatables.asp")	
				
	case "upload" 					: body=aspL.loadText("html/demo_asp/singleupload.resx")
		
	case "uploadfile" 				: aspL.exec("code/demo_asp/uploadfile.asp")	
	
	case "uploadmulti" 				: body=aspL.loadText("html/demo_asp/multiupload.resx")		
	
	'AJAX handlers
	'json.dump is very often used for AJAX handlers. "dump" basically means "flush" and "stop"	
	'All communication with the browser is done with JSON.
		
	case "sampleform1" 				: aspL.exec("code/demo_asp/sampleform1.asp")
	
	case "sampleform2" 				: aspL.exec("code/demo_asp/sampleform2.asp")
	
	case "sampleform3"				: aspL.exec("code/demo_asp/sampleform3.asp")
	
	case "sampleform4" 				: aspL.exec("code/demo_asp/sampleform4.asp")	
	
	case "sampleform5" 				: aspL.exec("code/demo_asp/sampleform5.asp")
	
	case "sampleform6" 				: aspL.exec("code/demo_asp/sampleform6.asp")	
	
	case "sampleform7" 				: aspL.exec("code/demo_asp/sampleform7.asp")
	
	case "sampleform8" 				: aspL.exec("code/demo_asp/sampleform8.asp")
	
	case "sampleform9" 				: aspL.exec("code/demo_asp/sampleform9.asp")
	
	case "sampleform10" 			: aspL.exec("code/demo_asp/sampleform10.asp")
	
	case "sampleform11" 			: aspL.exec("code/demo_asp/sampleform11.asp")

	case "sampleform12" 			: aspL.exec("code/demo_asp/sampleform12.asp")
	
	case "sampleform13" 			: aspL.exec("code/demo_asp/sampleform13.asp")	
	
	case "sampleform14" 			: aspL.exec("code/demo_asp/sampleform14.asp")	
	
	case "sampleform15" 			: aspL.exec("code/demo_asp/sampleform15.asp")	
	
	case "sampleform16" 			: aspL.exec("code/demo_asp/sampleform16.asp")	
	
	case "postdtt" 					: aspL.exec("code/demo_asp/postdtt.asp")	
	
	case "uploadfilejquery" 		: aspL.exec("code/demo_asp/uploadfile.asp") : aspL.die	''uploader
	
	case "welcome"					: aspL.exec("code/demo_asp/welcome.asp")
	
	case else 
				
		'get userfriendly url, if any (and launch a new handler-instance!)
		aspL.exec("code/demo_asp/404handler.asp")		
		
end select

%>