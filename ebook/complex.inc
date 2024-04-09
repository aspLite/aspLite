<%

dim form : set form=aspl.form
form.className="row"

dim name : set name=form.field("text")
name.add "name","name"
name.add "class","form-control"
name.add "label","Your name"
name.add "required",true
name.add "container","div"
name.add "containerclass","col-md-6"

dim years : set years=form.field("number")
years.add "name","years"
years.add "class","form-control"
years.add "label","For how many years are you an ASP developer?"
years.add "container","div"
years.add "containerclass","col-md-6"

dim birthdate : set birthdate=form.field("date")
birthdate.add "name","birthdate"
birthdate.add "class","form-control"
birthdate.add "label","Your birthdate"
birthdate.add "container","div"
birthdate.add "containerclass","col-md-6"

dim email : set email=form.field("email")
email.add "name","email"
email.add "class","form-control"
email.add "label","Your email"
email.add "container","div"
email.add "containerclass","col-md-6"

dim remarks : set remarks=form.field("textarea")
remarks.add "name","remarks"
remarks.add "class","form-control"
remarks.add "label","Remarks"
remarks.add "rows",5
remarks.add "container","div"
remarks.add "containerclass","col-md-12"

form.newline

dim submit : set submit=form.field("submit")
submit.add "html","Submit"
submit.add "class","btn btn-primary"
submit.add "style","margin-left:13px;width:100px"

dim reset : set reset=form.field("submit")
reset.add "html","Reset"
reset.add "class","btn btn-secondary"
reset.add "style","margin-left:7px;width:100px"

form.build
%>