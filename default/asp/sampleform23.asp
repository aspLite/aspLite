<%
'auto-refreshing form
'this can also be used as a "keep session alive" feature
'in that scenario you don't have to add fields, just build the form after form.reload=30

dim form : set form=aspl.form
form.listenTo "e","sampleform23"
form.target="sampleform23"
form.reload=30 'reload this form every XX seconds

'see global.asa for Application("visitors")
dim plainhtml : set plainhtml=form.field("plain")
plainhtml.add "html","<button type=""button"" onclick=""return false;"" class=""btn btn-primary"">Visitors online <span class=""badge badge-light"">" & Application("visitors") & "</span></button>"

form.build()

%>