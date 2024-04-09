<%
dim form : set form=aspl.form

dim plain : set plain=form.field("plain")
plain.add "html", "This adds plain/text or <u>HTML</u>. Check the console (f12)!"

dim script : set script=form.field("script")
script.add "text", "console.log('Add JavaScripts');"

form.build
%>