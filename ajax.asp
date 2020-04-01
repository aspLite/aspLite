<!-- #include file="asp/ASP.asp"-->
<%

select case lcase(asp.getrequest("myaction"))

	case "onload"		
	
		asp.flush "Hello world, Καλημέρα κόσμε (utf8-ready)"

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
		
		dim db,rs,field,json,counter,fieldvalue
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

	case "rssreader"
	
		dim rssreader
		set rssreader=asp.plugin("rssreader")
		asp.flush rssreader.load("http://rss.cnn.com/rss/cnn_topstories.rss")
		
			
	case else 'initial pageload!
	
		asp.flush asp.load("html/ajax.resx")

end select

%>