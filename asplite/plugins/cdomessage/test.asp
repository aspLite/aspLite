<%
dim mail: set mail=aspL.plugin("cdomessage")

mail.body="body text"
mail.subject="subject line"
mail.receiverEmail="abc@def.com"
mail.receiverName="Pieter Cooreman"
mail.fromname="ASP"
mail.fromemail="xx@xx.com"
'mail.send() 'commented out for security reasons - make sure to uncomment when you're ready

set mail=nothing
%>