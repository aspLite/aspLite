<%
'formbuilder sample - built on bootstrap css
aspl.exec("code/demo_asp/formbuilder/formbuilder.asp")
dim form : set form=new cls_formbuilder
form.targetDiv="body"

'form-submitted
if form.postback then

	'here you would typically perform additional validations, return error messages, and finally save to a database
	'below I also add a sample error & feedback message as comments to the form
	
	'error 
	dim errorM : set errorM=aspl.dict
	errorM.add "html","<h4>Example error message!</h4>"
	errorM.add "type","comment"
	errorM.add "tag","div"
	errorM.add "class","alert alert-danger"
	
	form.addField(errorM)
	
	'feedback
	dim feedback : set feedback=aspl.dict
	feedback.add "html","<h4>Successful submit!</h4>"
	feedback.add "type","comment"
	feedback.add "tag","div"
	feedback.add "class","alert alert-primary"
	
	form.addField(feedback)
	
	'rather than return the complete form in case of a successful submit, 
	'you can also build it here already. This will stop further exection
	'form.required=""
	'form.build()

end if	

'this hidden field is required in this demo, as the "e"-event is used in the event-handler
dim hidden : set hidden=aspl.dict
hidden.add "type","hidden"
hidden.add "name","e"
hidden.add "value","aspform"

form.addField(hidden)

'##########################################################################################

'a comment
dim comment : set comment=aspl.dict
comment.add "html","<h3>Ajax Form Builder</h3><hr><p class=""lead"">This ASP Ajax formbuilder makes creating Ajax forms in ASP very easy. The asp file ""code/demo_asp/aspform.asp"" returns the JSON data JavaScript (jQuery) needs to create a form.</p>"
comment.add "type","comment"
comment.add "tag","div"

form.addField(comment)

'##########################################################################################

dim firstname : set firstname=aspl.dict
firstname.add "label","First name:"
firstname.add "type","text"
firstname.add "name","firstname"
firstname.add "id","fname"
firstname.add "class","form-control"
firstname.add "maxlength",50
firstname.add "required",true

form.addField(firstname)

'##########################################################################################

dim lastname : : set lastname=aspl.dict
lastname.add "label","Last name:"
lastname.add "type","text"
lastname.add "name","lastname"
lastname.add "id","lname"
lastname.add "class","form-control"
lastname.add "maxlength",50
lastname.add "required",true

form.addField(lastname)

'##########################################################################################

dim email : set email=aspl.dict
email.add "label","Email:"
email.add "type","email"
email.add "name","email"
email.add "id","email"
email.add "value",""
email.add "class","form-control"

form.addField(email)

'##########################################################################################

dim aspyears : set aspyears=aspl.dict
aspyears.add "label","For how many years are you an ASP coder?"
aspyears.add "type","number"
aspyears.add "id","aspyears"
aspyears.add "name","aspyears"
aspyears.add "class","form-control"

form.addField(aspyears)

'##########################################################################################

dim yesno : set yesno=aspl.dict
yesno.add "label","Are you still coding ASP?"
yesno.add "type","select"
yesno.add "id","yesno"
yesno.add "name","yesno"
yesno.add "class","form-control"
'in case of multiple options (like selectboxes, radiobuttons, etc), use an array of arrays
yesno.add "options",Array(Array("","Please select"),Array("true","Yes"), Array("false","No"), Array("dunno","Don't know"))

form.addField(yesno)

'##########################################################################################

dim radio : set radio=aspl.dict
radio.add "label","How would you rate yourself as a developer?"
radio.add "type","radio"
radio.add "id","radio"
radio.add "name","radio"
'in case of multiple options (like selectboxes, radiobuttons, etc), use an array of arrays
radio.add "options",Array(Array(0,"Beginner"),Array(1,"Intermediate"), Array(2,"Advanced"))

form.addField(radio)

'##########################################################################################
'let's recycle the country list we already use in the DataTables sample

aspl.exec("code/demo_asp/datatables/includes.asp")
dim countrylist : set countrylist=new cls_countrylist
set countrylist=countrylist.list

dim countries : set countries=aspl.dict
countries.add "label","Where do you live?"
countries.add "type","select"
countries.add "id","countries"
countries.add "name","countries"
countries.add "class","form-control"
'VBScript dictionary (key: (option)value, pair: (option)text)
countries.add "options",countrylist 

form.addField(countries)

'##########################################################################################
'let's add a recordset directly as well

dim db : set db = aspl.plugin("database") : db.path="db/sample.mdb"
'always alias to keyV & pairV
'ADO recordset rs(0): keyV | rs(1): pairV
set rs=db.execute("select iId as keyV, sName as pairV from person order by sName asc")

dim persons : set persons=aspl.dict
persons.add "label","Who is your favorite?"
persons.add "type","select"
persons.add "id","persons"
persons.add "name","persons"
persons.add "class","form-control"
persons.add "options",rs

form.addField(persons)

'##########################################################################################

dim usercomments : set usercomments=aspl.dict
usercomments.add "label","Additional user comments (if any)"
usercomments.add "type","textarea"
usercomments.add "name","usercomments"
usercomments.add "id","usercomments"
usercomments.add "rows",3
usercomments.add "class","form-control"

form.addField(usercomments)

'##########################################################################################
' another comment

dim anothercomment : set anothercomment=aspl.dict
anothercomment.add "html","Yet another comment."
anothercomment.add "type","comment"
anothercomment.add "tag","div"
anothercomment.add "class","alert alert-warning"
anothercomment.add "style","margin-top:20px"

form.addField(anothercomment)

'##########################################################################################

dim birthdate : set birthdate=aspl.dict
birthdate.add "label","Birthday:"
birthdate.add "type","date"
birthdate.add "name","birthdate"
birthdate.add "id","bDate"
birthdate.add "class","form-control"

form.addField(birthdate)

'##########################################################################################

dim submit : set submit=aspl.dict
submit.add "label","<br><br>"
submit.add "type","submit"
submit.add "name","btnAction"
submit.add "id","btnAction"
submit.add "value","Submit"
submit.add "class","btn btn-primary"

form.addField(submit)

'##########################################################################################

form.build()

%>