<%
set form=aspl.form

set canvas = form.field ("element")
canvas.add "tag","canvas"
canvas.add "id","demo"
canvas.add "style","background-color:#111"

set field = form.field ("plain")
field.add "html", aspl.loadText("default/html/canvas.resx")

form.build
%>