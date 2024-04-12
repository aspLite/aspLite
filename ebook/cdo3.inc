<%

dim cdomessage : set cdomessage=aspL.plugin("cdomessage")	
cdomessage.smtpserver="localhost"
cdomessage.smtpport=25
cdomessage.sendusing=1
cdomessage.pickupdir="C:\inetpub\mailroot\Pickup"
cdomessage.fromname="aspLite"
cdomessage.fromemail="info@iseral.be"
cdomessage.receiveremail="pietercooreman@gmail.com"
cdomessage.subject="aspLite testmail"	

dim body : body=aspl.loadtext("ebook/mail.txt")
body=replace(body,"[systemname]","Setlist Planner",1,-1,1)
body=replace(body,"[body]","Welcome to Setlist Planner",1,-1,1)

cdomessage.body=body
'cdomessage.send
set cdomessage=nothing

%>