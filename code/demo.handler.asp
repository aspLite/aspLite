<%
on error resume next

dim html,titletag,body

titletag="aspLite demo"

'load the template of the demosite
html=aspL.loadText("html/demo_asp/demo.resx")

'create a json object for all AJAX communication with the client
'be careful, this object will be used all over this demo site! do not destroy it!
dim json : set json=aspL.plugin("json")

'here you can typically add some sort of eventhandler (what did the user do exactly?)
'in general: if the code for dealing with the event is not large, include it here.
'if the code is getting complex, use yet another codebehind file, like in the examples below (helloworld, mail, json, etc)

select case lcase(aspL.getRequest("e")) '"event"
		
	case "clicklink" : body="Link was clicked"		
		
	case "clickbutton" : body="Regular form was submitted"		
		
	case "loadclass"
	
		'CONDITIONAL load of asp page = include file
		aspL.exec("code/demo_asp/class.asp")	
		
		dim testObj
		set testObj=new cls_test
		body=testObj.hello
		set testObj=nothing	
		
	
	case "downloadlargefile" : aspL.dumpBinary("html/demo_asp/largefile.jpg")
	
	case "downloadsmallfile" : aspL.dumpBinary("html/demo_asp/smallfile.jpg")
	
	case "helloworld" : aspL.exec("code/demo_asp/helloworld.asp")	
		
	case "randomizer" : aspL.exec("code/demo_asp/randomtext.asp")		
		
	case "db" : aspL.exec("code/demo_asp/datasample.asp")	
		
	case "datatables" : aspL.exec("code/demo_asp/datatables.asp")
	
	case "datatables_ssp" : body=aspL.loadText("html/demo_asp/datatables_ssp.resx")	

	case "json_datatables_data" : aspL.exec("code/demo_asp/datatables_ssp.asp")	
	
	case "json_datatables" : aspL.exec("code/demo_asp/json_datatables.asp")
		
	case "ckeditor" : aspL.exec("code/demo_asp/ckeditor.asp")
	
	case "ckeditor5" : aspL.exec("code/demo_asp/ckeditor5.asp")
				
	case "upload" : body=aspL.loadText("html/demo_asp/singleupload.resx")
		
	case "uploadfile" : aspL.exec("code/demo_asp/uploadfile.asp")	
	
	case "uploadmulti" : body=aspL.loadText("html/demo_asp/multiupload.resx")		
	
	case "jszip" : aspL.exec("code/demo_asp/jszip.asp")	
		
	case "jspdf" : body=aspL.loadText("html/demo_asp/jspdf.resx")
	
	case "codemirror" : aspL.exec("code/demo_asp/codemirror.asp")	
	
	case "jqueryui" : aspL.exec("code/demo_asp/jqueryui.asp")
		
	case "login" : html=aspL.loadText("html/demo_asp/login.resx")	
		'this will load a completely different template because we override the html variable

		
	'AJAX handlers
	'json.dump is very often used for AJAX handlers. "dump" basically means "flush" and "stop"	
	'All communication with the browser is done with JSON.
		
	case "ajaxhello" : json.dump("Hello " & aspL.sanitize(aspL.getRequest("yourname")) & ".<br><br>Hashed (sha256):<br><br><textarea class=""form-control"">" & aspL.plugin("sha256").sha256(aspL.getRequest("yourname")) & "</textarea>")

	case "submit1" : json.dump("Form 1 save button was clicked - " & aspL.getrequest("yourname"))
		
	case "linksubmit" : json.dump("Form 1 submitted by link - " & aspL.getrequest("yourname"))
	
	case "delete1" : json.dump("Form 1 delete button was clicked - " & aspL.getrequest("yourname"))
		
	case "submit2" : json.dump("Form 2 save button was clicked")
	
	case "delete2" : json.dump("Form 2 delete button was clicked")
		
	case "buttonclick" : json.dump("Regular button was clicked")
		
	case "clicklink1" : json.dump("First link clicked...")
	
	case "clicklink2" : json.dump("Second link clicked...")
	
	case "returnbool" : json.dump(aspl.convertStr(aspL.plugin("randomizer").randomnumber(0,100)>49))
		
	case "returndata" : aspL.exec("code/demo_asp/datatable.asp")	
	
	case "hash" : aspL.exec("code/demo_asp/hash.asp")	
		
	case "json" : aspL.exec("code/demo_asp/json.asp")
	
	case "json2html" : aspL.exec("code/demo_asp/json.asp")
	
	case "postdtt" : aspL.exec("code/demo_asp/postdtt.asp")
		
	case "sendmail" : aspL.exec("code/demo_asp/mail.asp")		
	
	case "atom" : json.dump(aspL.plugin("atom").read("https://github.com/timeline"))

	case "rss" : json.dump(aspL.plugin("rss").read("http://rss.cnn.com/rss/cnn_topstories.rss"))
		
	case "jpg" : aspL.exec("code/demo_asp/jpg.asp")	
	
	case "uploadjquery" : json.dump(aspL.loadText("html/demo_asp/uploadjquery.resx"))
	
	case "uploadfilejquery" : aspL.exec("code/demo_asp/uploadfile.asp") : aspL.die	''uploader
	
	case "rate" : aspL.exec("code/demo_asp/rate.asp")	
			
	case "onload" :	json.dump("Hello world, Καλημέρα κόσμε (utf8-ready)")		
		
	case else body="No (known) action was detected. Initial load." 'default content for REGULAR handler
			
		'get userfriendly url, if any (and launch a new handler-instance!)
		aspL.exec("code/demo_asp/404handler.asp")		
		
end select

on error goto 0

%>