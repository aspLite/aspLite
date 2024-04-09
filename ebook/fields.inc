<%

dim form : set form=aspl.form

if form.postback then

	form.write "Hello " & aspl.htmlEncode(aspl.getRequest("name"))
	form.newline
	
	form.write "You're an ASP developer for  " & aspl.htmlEncode(aspl.getRequest("years")) & " years."
	form.newline
	
	form.write "Your birthdate: " & aspl.htmlEncode(aspl.getRequest("birthdate"))
	form.newline	
	
	form.write "Your email: " & aspl.htmlEncode(aspl.getRequest("email"))
	form.newline	
	
	form.write "Your remarks: " & aspl.htmlEncode(aspl.getRequest("remarks"))
	form.newline
	
	form.write "Hidden field: " & aspl.htmlEncode(aspl.getRequest("hidden"))
	
	form.build

end if

dim hidden : set hidden=form.field("hidden")
hidden.add "name","hidden"
hidden.add "value","12345"

dim name : set name=form.field("text")
name.add "name","name"
name.add "class","form-control"
name.add "label","Your name"
name.add "required",true

form.newline

dim years : set years=form.field("number")
years.add "name","years"
years.add "class","form-control"
years.add "label","For how many years are you an ASP developer?"

form.newline

dim birthdate : set birthdate=form.field("date")
birthdate.add "name","birthdate"
birthdate.add "class","form-control"
birthdate.add "label","Your birthdate"

form.newline

dim email : set email=form.field("email")
email.add "name","email"
email.add "class","form-control"
email.add "label","Your email"

form.newline

dim remarks : set remarks=form.field("textarea")
remarks.add "name","remarks"
remarks.add "class","form-control"
remarks.add "label","Remarks"
remarks.add "rows",5

form.newline

dim submit : set submit=form.field("submit")
submit.add "html","Submit"
submit.add "class","btn btn-primary"

dim reset : set reset=form.field("reset")
reset.add "html","Reset"
reset.add "class","btn btn-secondary"

form.build
%>