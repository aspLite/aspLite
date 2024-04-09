<%

dim form : set form=aspl.form

if form.postback then

	form.writejs(form.redirect("redirectNew","redirect","message=" & server.urlencode("Hello there!")))

end if

dim redirect : set redirect=form.field("submit")
redirect.add "html","Redirect"
redirect.add "class","btn btn-primary"

form.build

%>