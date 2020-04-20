<%
dim mail: set mail=aspL.plugin("cdomessage")

mail.body="<pre>" & aspL.loadText("html/demo_asp/mail.txt") & "</pre>"
mail.subject="ABCDEFGHIJKLMNOPQRSTUVWXYZ /0123456789  abcdefghijklmnopqrstuvwxyz £©µÀÆÖÞßéöÿ  –—‘“”„†•…‰™œŠŸž€ ΑΒΓΔΩαβγδω АБВГДабвгд  ∀∂∈ℝ∧∪≡∞ ↑↗↨↻⇣ ┐┼╔╘░►☺♀ ﬁ�⑀₂ἠḂӥẄɐː⍎אԱა"
mail.receiverEmail="pietercooreman@gmail.com"
mail.receiverName="Pieter Cooreman"
mail.fromname="ASP"
mail.fromemail="info@quickersite.com"
'mail.send() 'commented out for security reasons - make sure to uncomment when you're ready

set mail=nothing

aspL.dump "Mail sent"
%>