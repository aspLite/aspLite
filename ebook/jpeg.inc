<%
dim form : set form=aspl.form

dim jpg : set jpg=aspL.plugin("jpg")
jpg.maxsize=140 'px - max: 2560px
jpg.path="/default/html/smallfile.jpg"

form.write "These are the 4 available effects:"
form.newline

dim i
for i=0 to 3
	jpg.effect=i
	form.write "<img style=""margin:2px;border-radius:5px;"" src=""" & jpg.src & """ />"
next

form.newline

form.write "The next 4 photo's will be cropped to squares:"
form.newline

jpg.square=1 
jpg.maxsize=140
for i=0 to 3
	jpg.effect=i
	form.write "<img style=""margin:2px;border-radius:5px;"" src=""" & jpg.src & """ />"
next

form.build
%>