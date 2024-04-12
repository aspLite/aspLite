<%
class cls_asplite_cdomessage

	public receiverEmail,receiverName, subject, body, attachments, fromname, fromemail
	public smtpusessl, smtpusername, smtpserver, smtpport, sendusing, pickupdir, smtpuserpw, copyto
	
	private sub class_initialize	
		set attachments=aspL.dict	
	end sub
	
	private sub class_terminate
		set attachments=nothing
	end sub
	
	Public function send
	
		On Error Resume next
		
		body=wrapInHTML(prepareforEmail(body),subject)
		body=replace(body,">" & vbcrlf &".",">.",1,-1,1)
	
		dim myMail, fileKey	
		set myMail=Server.CreateObject("CDO.Message")
		
		'configuration object
		myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")=smtpserver
		myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=aspl.convertNmbr(smtpport)
		myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout")=60
		myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing")=aspl.convertNmbr(sendusing)
		myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = pickupdir

		if smtpusername<>"" then
			myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")=1
			myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername")=smtpusername
			myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword")=smtpuserpw
		end if
		
		if aspl.convertBool(smtpusessl) then
			myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
		else
			myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
		end if

		myMail.Configuration.Fields.Update
		myMail.Bodypart.ContentMediaType = "text/html"
		myMail.Bodypart.Charset = "utf-8"
		myMail.Bodypart.ContentTransferEncoding = "quoted-printable"
		myMail.Subject=subject
		myMail.To=receiverName&"<"&receiverEmail&">"
		myMail.From=fromname & "<"& fromemail &">"
		myMail.ReplyTo=fromemail
		
		'attachments
		for each fileKey in attachments
			myMail.AddAttachment attachments(fileKey)
		next 
		
		myMail.HtmlBody=body
		myMail.HTMLBodypart.Charset = "utf-8"
		myMail.send
		
		if not aspl.isEmpty(copyto) then
			body=replace(body,"</body>","<table border=""0"" style=""border-style:none;width:100%"" cellpadding=""10""><tr><td style=""font-family:Arial;font-size:16px;background-color:silver"">Copy of email to " & receiverName & " - " & receiverEmail & "</td></tr></table></body>",1,-1,1)
			myMail.HtmlBody=body
			myMail.Subject=subject & " (copy)"
			myMail.To=copyto
			myMail.send
		end if
		
		set MyMail=nothing	
			
		aspl.asperror("Something goes wrong when sending emails. See details below:")		
		
		On Error Goto 0
		
	end function
	
	function wrapInHTML(body,subject)
	
		if instr(1,body,"<body",vbTextCompare)=0 then
			
			'assuming no complete html document was used, so we complete it
			
			wrapInHTML="<!doctype html>" & vbcrlf
			wrapInHTML=wrapInHTML & "<html>" & vbcrlf & "<head>" & vbcrlf 
			wrapInHTML=wrapInHTML & "<meta charset=""utf-8"">"  & vbcrlf 
			wrapInHTML=wrapInHTML & "<meta name=""viewport"" content=""width=device-width, initial-scale=1, minimum-scale=1"">"  & vbcrlf 
			wrapInHTML=wrapInHTML & "<title>"& subject &"</title>" & vbcrlf & "</head>" & vbcrlf 
			wrapInHTML=wrapInHTML& "<body>" & vbcrlf & body & vbcrlf & "</body>" & vbcrlf & "</html>" & vbcrlf 
		
		else
		
			wrapInHTML=body
		
		end if
		
	end function
	
	
	function prepareforEmail(value)
	
		prepareforEmail=value
		
		if left(trim(prepareforEmail),1)<>"<" then
			prepareforEmail="<div>" & prepareforEmail & "</div>"
		end if		
				
	end function
	
end class

%>
