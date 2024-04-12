<%

dim cdomessage : set cdomessage=aspL.plugin("cdomessage")	
cdomessage.smtpserver="smtp-auth.mailprotect.be"
cdomessage.smtpport=465
cdomessage.sendusing=2
cdomessage.pickupdir="C:\inetpub\mailroot\Pickup"
cdomessage.fromname="Setlist Planner"
cdomessage.fromemail="info@setlist.be"
cdomessage.receiveremail="pietercooreman@gmail.com"
cdomessage.smtpusername="info@setlist.be"
cdomessage.smtpuserpw="xxxxx"
cdomessage.smtpusessl=true
cdomessage.subject="aspLite testmail"	
cdomessage.body="aspLite testmail"
cdomessage.attachments.add "Attachment",server.mappath("ebook/attachment.rtf")
'cdomessage.send
set cdomessage=nothing
	
%>