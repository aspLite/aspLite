<%

dim form : set form=aspl.form

if form.postback then

	form.writejs(form.redirect("redirectNewMultiple","redirect1","message=Hello%20Redirect1"))
	form.writejs(form.redirect("redirectNewMultiple","redirect2","message=Hello%20Redirect2"))
	form.writejs(form.redirect("redirectNewMultiple","redirect3","message=Hello%20Redirect3"))
	
	form.write "Ok, we're done here."
	
	form.build

end if

dim redirect : set redirect=form.field("submit")
redirect.add "html","Redirect"
redirect.add "class","btn btn-primary"

form.build

%>