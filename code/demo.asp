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
select case lcase(asp.getRequest("e")) '"event"
		
	case "clicklink"
		
		body="Link was clicked"		
		
	case "clickbutton"
	
		body="Regular form was submitted"		
		
	case "loadclass"
	
		'CONDITIONAL load of asp page = include file
		asp.exec("code/includes/class.asp")	
		
		dim testObj
		set testObj=new cls_test
		body=testObj.hello
		set testObj=nothing	
		
	
	case "downloadlargefile"
	
		asp.flushBinaryFile("html/largefile.jpg")
	
	case "downloadsmallfile"
	
		asp.flushBinaryFile("html/smallfile.jpg")
	
	case "helloworld"
	
		'hello world plugin example
		dim helloworld
		set helloworld=asp.plugin("helloworld")
		body=helloworld.hw()
		'or shorter
		titletag=asp.plugin("helloworld").hw()
		set helloworld=nothing
		
	case "randomizer"
	
		'randomizer plugin example
		dim i, randomizer
		set randomizer=asp.plugin("randomizer")

		body="ASP randomizer-plugin: "
		for i=1 to 25
			'generate some random words with random lengths (5-10)
			body=body & randomizer.randomtext(randomizer.randomnumber(5,10)) & " "
		next	
		
		
	case "accessdb"
	
		'database plugin example
		set db=asp.plugin("accessDB")
		db.path="db/sample.mdb"

		body=body & "Access database-plugin:<ul> " 

		set rs=db.execute("select * from person")

		while not rs.eof
			body=body & "<li>" & rs("sName") & "</li>"
			rs.movenext
		wend 

		body=body & "</ul>"
	

	'AJAX handlers
		
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
	
		'get the bootstrap thumbnail template
		dim bsTemplate
		bsTemplate=asp.load("html/bootstrapthumb.txt")

		dim jpg
		set jpg=asp.plugin("jpg")
		jpg.maxsize=600 'px - max: 2560px
		
		'this looks more complex than it is, as this sample is supposed to work in various setups
		'by default, this would rather look like jpg.path="/images/img.jpg" where this path is relative to your directory
		jpg.path=replace(request.servervariables("path_info"),"demo.asp","",1,-1,1) & asp_path & "/plugins/jpg/sample.jpg"
		
		dim specialeffects
		specialeffects=replace(bsTemplate,"[src]",jpg.src,1,-1,1)
		specialeffects=replace(specialeffects,"[caption]","Resizing",1,-1,1)
		
		'color effects
		jpg.effect=1 'b/w
			specialeffects=specialeffects & replace(bsTemplate,"[src]",jpg.src,1,-1,1)
			specialeffects=replace(specialeffects,"[caption]","black/white",1,-1,1)
		jpg.effect=2 'gray
			specialeffects=specialeffects & replace(bsTemplate,"[src]",jpg.src,1,-1,1)
			specialeffects=replace(specialeffects,"[caption]","gray",1,-1,1)
		jpg.effect=3 'sepia
			specialeffects=specialeffects & replace(bsTemplate,"[src]",jpg.src,1,-1,1)
			specialeffects=replace(specialeffects,"[caption]","sepia",1,-1,1)
		
		'crop to rectangle
		jpg.effect=0
		jpg.fsr=1 
			specialeffects=specialeffects & replace(bsTemplate,"[src]",jpg.src,1,-1,1)
			specialeffects=replace(specialeffects,"[caption]","rectangle",1,-1,1)
		
		asp.flush "<div class=""card-columns"">" & specialeffects & "</div>"	
	
	
	case "onload" 'default content for AJAX handler	
	
		asp.flush "Hello world, Καλημέρα κόσμε (utf8-ready)"
		
	case else 'default content for REGULAR handler
	
		body="No (known) action was detected. Initial load."
		
end select

%>