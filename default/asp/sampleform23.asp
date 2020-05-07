<%
'auto-refreshing form
dim form : set form=aspl.form
form.listenTo "e","sampleform23"
form.target="sampleform23"
form.reload=30 'reload this form every XX seconds

'see global.asa for Application("visitors")
dim plainhtml : set plainhtml=form.field("plain")
plainhtml.add "html","<button type=""button"" onclick=""return false;"" class=""btn btn-primary"">Visitors online <span class=""badge badge-light"">" & Application("visitors") & "</span></button>"

form.build()

%>