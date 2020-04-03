<%
'###########################################################################
'## include these 2 lines to start each codebehind file for security reasons
option explicit  
asp.codebehind() 
'###########################################################################

dim html,titletag,body
dim db,rs,counter,json,field,fieldvalue

titletag="aspLite demo"

'load the template
'resx files are never served to browsers, so they are safer to use
html=asp.load("html/demo.resx")

'here you can typically add some sort of eventhandler (what exactly was clicked on?)
select case lcase(asp.getRequest("action"))
		
	case "clicklink"
		
		body="<p>Link was clicked</p>"		
		
	case "clickbutton"
	
		body="<p>Button was clicked</p>"		
		
	case "loadclass"
	
		'CONDITIONAL load of asp page = include file
		asp.exec("code/includes/class.asp")	
		
		dim testObj
		set testObj=new cls_test
		body="<p>" & testObj.hello & "</p>"
		set testObj=nothing	
		
	
	case "downloadlargefile"
	
		asp.flushBinaryFile("html/largefile.jpg")
	
	case "downloadsmallfile"
	
		asp.flushBinaryFile("html/smallfile.jpg")
	
	case "helloworld"
	
		'hello world plugin example
		dim helloworld
		set helloworld=asp.plugin("helloworld")
		body="<p>" & helloworld.hw() & "</p>"
		'or shorter
		titletag=asp.plugin("helloworld").hw()
		set helloworld=nothing
		
	case "randomizer"
	
		'randomizer plugin example
		dim i, randomizer
		set randomizer=asp.plugin("randomizer")

		body="<p>ASP randomizer-plugin: "
		for i=1 to 20
			'generate some random words with random lengths (5-10)
			body=body & randomizer.randomtext(randomizer.randomnumber(5,10)) & " "
		next	
		body=body & "</p>"
		
	case "accessdb"
	
		'database plugin example
		set db=asp.plugin("accessDB")
		db.path="db/sample.mdb"

		body=body & "<p>Access database-plugin: " 

		set rs=db.execute("select * from person")

		while not rs.eof
			body=body & rs("sName") & " "
			rs.movenext
		wend 

		body=body & "</p>"
		
	case else
	
		body="<p>No (known) action was detected. Initial load.</p>"	
	

end select

'AJAX handlers

select case lcase(asp.getrequest("myaction"))

	case "onload"		
	
		asp.flush "Hello world, Καλημέρα κόσμε (utf8-ready)"
		
	case "ajaxhello"
		
		asp.flush "Hello " & asp.sanitize(asp.URLDecode(asp.getRequest("yourname")))

	case "submit1"
	
		asp.flush "Form 1 save button was clicked - " & asp.getrequest("yourname")
		
	case "linksubmit"
	
		asp.flush "Form 1 submitted by link - " & asp.getrequest("yourname")
	
	case "delete1"
	
		asp.flush "Form 1 delete button was clicked - " & asp.getrequest("yourname")
		
	case "submit2"
	
		asp.flush "Form 2 save button was clicked"
	
	case "delete2"
	
		asp.flush "Form 2 delete button was clicked"
		
	case "buttonclick"
	
		asp.flush "Regular button was clicked"
		
	case "clicklink1"

		asp.flush "First link clicked..."
	
	case "clicklink2"

		asp.flush "Second link clicked..."	
	
	case "returnbool"
	
		'returns a random boolean to the browser		
		asp.flush asp.plugin("randomizer").randomnumber(0,100)>49	
		
	case "returnjsondata"

		counter=0
		
		set db=asp.plugin("accessdb")
		db.path="db/sample.mdb"
		
		set rs=db.GetDynamicRS

		rs.open ("select * from person")
		
		set json=asp.plugin("json")
		json.data.Add "persons", json.Collection()	
		
		while not rs.eof

			json.data("persons").Add counter,json.Collection()
			
			for each field in rs.fields	
				fieldvalue=rs(field.name)
				json.data("persons").item(counter).add asp.sanitize(field.name),asp.sanitize(fieldvalue)
			next
			
			counter=counter+1
			
			rs.movenext
		
		wend
		
		asp.flush json.JSONoutput
		
	case "sendmail"
	
		dim mail
		set mail=asp.plugin("cdomessage")
		mail.body="<pre>" & asp.load("html/utf8.txt") & "</pre>"
		mail.subject="ABCDEFGHIJKLMNOPQRSTUVWXYZ /0123456789  abcdefghijklmnopqrstuvwxyz £©µÀÆÖÞßéöÿ  –—‘“”„†•…‰™œŠŸž€ ΑΒΓΔΩαβγδω АБВГДабвгд  ∀∂∈ℝ∧∪≡∞ ↑↗↨↻⇣ ┐┼╔╘░►☺♀ ﬁ�⑀₂ἠḂӥẄɐː⍎אԱა"
		mail.receiverEmail="pietercooreman@gmail.com"
		mail.receiverName="Pieter Cooreman"
		mail.fromname="ASP"
		mail.fromemail="info@quickersite.com"
		'mail.send() 'commented out for security reasons - make sure to uncomment when you're ready
		set mail=nothing
		
		asp.flush "Mail sent"

	case "rss"
	
		dim rss
		set rss=asp.plugin("rss")
		asp.flush rss.read("http://rss.cnn.com/rss/cnn_topstories.rss")
		
	case "jpg"	

		dim jpg
		set jpg=asp.plugin("jpg")
		jpg.maxsize=500 'px - max: 2560px
		
		'this looks more complex than it is, as this sample is supposed to work in various setups
		'by default, this would rather look like jpg.path="/images/img.jpg" where this path is relative to your directory
		jpg.path=replace(request.servervariables("path_info"),"demo.asp","",1,-1,1) & asp_path & "/plugins/jpg/sample.jpg"
		
		dim specialeffects
		specialeffects="normal resize:<br><img class=""img-fluid"" src=""" & jpg.src & """ /><br>"
		
		'color effects
		jpg.effect=1 'b/w
		specialeffects=specialeffects & "black/white<br><img class=""img-fluid"" src=""" & jpg.src & """ /><br>"
		jpg.effect=2 'gray
		specialeffects=specialeffects & "gray<br><img class=""img-fluid"" src=""" & jpg.src & """ /><br>"
		jpg.effect=3 'sepia
		specialeffects=specialeffects & "sepia<br><img class=""img-fluid"" src=""" & jpg.src & """ /><br>"
		
		'crop to rectangle
		jpg.effect=0
		jpg.fsr=1 
		specialeffects=specialeffects & "rectangle:<br><img class=""img-fluid"" src=""" & jpg.src & """ />"	
		
		asp.flush specialeffects		
		

end select




%>