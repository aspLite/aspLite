<%
'formbuilder sample - built on bootstrap css
aspl.exec("code/demo_asp/formbuilder/formbuilder.asp")

dim form : set form=new cls_formbuilder
form.listenTo "e","sampleform5"
form.targetDiv="sampleform5"
form.id="sampleForm"

'form-submitted
if form.postback then

	form.doScroll=true
	
	'here you would typically perform additional validations, return error messages, and finally save to a database (or delete)
	'below I add a feedback message as a comment to the form
		
	'feedback
	dim feedback : set feedback=form.field
	
	select case aspl.getRequest("aspFormAction")
	
		case "save"
		
			feedback.add "html","Successfully saved!"
			feedback.add "type","comment"
			feedback.add "tag","div"
			feedback.add "class","alert alert-primary"			
		
		case "delete"
	
			feedback.add "html","Record Deleted!"
			feedback.add "type","comment"
			feedback.add "tag","div"
			feedback.add "class","alert alert-warning"
			
			'in case of a delete, we may want to no longer show anything but this feedback message
			form.build()
	
	end select
	
	'rather than return the complete form in case of a successful submit, 
	'you can also build it here already. This will stop further exection
	'form.build()

end if	

'##########################################################################################

dim aspFormAction : set aspFormAction=form.field
aspFormAction.add "type","hidden"
aspFormAction.add "name","aspFormAction"
aspFormAction.add "id","aspFormAction"

'##########################################################################################

dim email : set email=form.field
email.add "label","Email:"
email.add "type","email"
email.add "name","email"
email.add "id","email"
email.add "value",""
email.add "class","form-control"

'##########################################################################################

dim aspyears : set aspyears=form.field
aspyears.add "label","For how many years are you an ASP coder?"
aspyears.add "type","number"
aspyears.add "id","aspyears"
aspyears.add "name","aspyears"
aspyears.add "class","form-control"

'##########################################################################################

dim yesno : set yesno=form.field
yesno.add "label","Are you still coding ASP?"
yesno.add "type","select"
yesno.add "id","yesno"
yesno.add "name","yesno"
yesno.add "class","form-control"
'in case of multiple options (like selectboxes, radiobuttons, etc), use an array of arrays
yesno.add "options",Array(Array("","Please select"),Array("true","Yes"), Array("false","No"), Array("dunno","Don't know"))

'##########################################################################################

dim radio : set radio=form.field
radio.add "label","How would you rate yourself as a developer?"
radio.add "type","radio"
radio.add "id","radio"
radio.add "name","radio"
'in case of multiple options (like selectboxes, radiobuttons, etc), use an array of arrays
radio.add "options",Array(Array(0,"Beginner"),Array(1,"Intermediate"), Array(2,"Advanced"))

'##########################################################################################
'let's recycle the country list we already use in the DataTables sample

aspl.exec("code/demo_asp/datatables/includes.asp")
dim countrylist : set countrylist=new cls_countrylist
set countrylist=countrylist.list

dim countries : set countries=form.field
countries.add "label","Where do you live?"
countries.add "type","select"
countries.add "id","countries"
countries.add "name","countries"
countries.add "class","form-control"
'VBScript dictionary (key: (option)value, pair: (option)text)
countries.add "options",countrylist 

'##########################################################################################
'let's add a recordset directly as well
'always alias to keyV & pairV
'ADO recordset rs(0): keyV | rs(1): pairV
set rs=db.execute("select iId as keyV, sName as pairV from person order by sName asc")

dim persons : set persons=form.field
persons.add "label","Who is your favorite?"
persons.add "type","select"
persons.add "id","persons"
persons.add "name","persons"
persons.add "class","form-control"
persons.add "options",rs

'##########################################################################################

dim submit : set submit=form.field
submit.add "type","submit"
submit.add "name","btnSave"
submit.add "value","Submit"
submit.add "style","margin-top:15px"
submit.add "class","btn btn-primary"
submit.add "container","span"
submit.add "onclick","$('#aspFormAction').val('save')"

dim delete : set delete=form.field
delete.add "type","submit"
delete.add "name","btnDelete"
delete.add "value","Delete"
delete.add "style","margin-top:15px"
delete.add "class","btn btn-danger"
delete.add "container","span"
delete.add "containerstyle","margin-left:10px"
delete.add "onclick","if (confirm('Are you sure?')) {$('#aspFormAction').val('delete')} else {return false}"

'##########################################################################################

form.build()

%>