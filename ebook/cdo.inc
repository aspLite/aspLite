<%

dim form : set form=aspl.form

if form.postback then

	if aspl.CheckEmailSyntax(aspl.getRequest("email")) then

		dim cdomessage : set cdomessage=aspL.plugin("cdomessage")	
		cdomessage.smtpserver="localhost"
		cdomessage.smtpport=25
		cdomessage.sendusing=1
		cdomessage.pickupdir="C:\inetpub\mailroot\Pickup"
		cdomessage.fromname="aspLite"
		cdomessage.fromemail="info@iseral.be"
		cdomessage.receiveremail=aspl.getRequest("email")
		cdomessage.subject="aspLite testmail"	
		cdomessage.body="aspLite testmail"
		'cdomessage.send
		set cdomessage=nothing
		
		form.alert "success","Mail sent!"
		
		form.build
	
	else
	
		form.alert "danger","Incorrect email address!"
	
	end if

end if

dim email : set email=form.field("email")
email.add "required",true
email.add "name","email"
email.add "class","form-control"

form.newline

dim send : set send=form.field("submit")
send.add "html","Send email"
send.add "class","btn btn-info"

form.build

%>