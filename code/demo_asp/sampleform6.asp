<%

dim form : set form=new cls_formbuilder
form.listenTo "e","sampleform6"
form.targetDiv="sampleform6"
form.id="multiButtons"

'feedback
dim feedback : set feedback=form.field

'form-submitted
if form.postback then

	form.doScroll=true
	
	'here you would typically perform additional validations, return error messages, and finally save to a database (or delete)
	'below I add a feedback message as a comment to the form
		
	select case aspl.getRequest("mb_aspFormAction")
	
		case "save"
		
			feedback.add "html","Save button was clicked"
			feedback.add "type","comment"
			feedback.add "tag","div"
			feedback.add "class","alert alert-primary"			
		
		case "delete"
	
			feedback.add "html","Delete button was clicked"
			feedback.add "type","comment"
			feedback.add "tag","div"
			feedback.add "class","alert alert-danger"
			
		case "link"
	
			feedback.add "html","Link button clicked"
			feedback.add "type","comment"
			feedback.add "tag","div"
			feedback.add "class","alert alert-info"
		
		case "button"
	
			feedback.add "html","Regular button clicked"
			feedback.add "type","comment"
			feedback.add "tag","div"
			feedback.add "class","alert alert-warning"

	
	end select
	
	'rather than return the complete form in case of a successful submit, 
	'you can also build it here already. This will stop further exection
	'form.build()

else

	'form is not in "postback"-state yet.
	'let's reuse the feedback field to return an initial intro message

	feedback.add "html","Form is not yet submitted"
	feedback.add "type","comment"
	feedback.add "tag","div"
	feedback.add "class","alert alert-dark"

end if	

'##########################################################################################

dim aspFormAction : set aspFormAction=form.field
aspFormAction.add "type","hidden"
aspFormAction.add "name","mb_aspFormAction"
aspFormAction.add "id","mb_aspFormAction"

'##########################################################################################

dim submit : set submit=form.field
submit.add "type","submit"
submit.add "value","Save"
submit.add "style","margin-top:15px"
submit.add "class","btn btn-primary"
submit.add "container","span"
submit.add "onclick","$('#mb_aspFormAction').val('save')"

dim delete : set delete=form.field
delete.add "type","submit"
delete.add "value","Delete"
delete.add "style","margin-top:15px"
delete.add "class","btn btn-danger"
delete.add "container","span"
delete.add "containerstyle","margin-left:10px"
delete.add "onclick","$('#mb_aspFormAction').val('delete');"

dim link : set link=form.field
link.add "type","comment"
link.add "tag","a"
link.add "html","Link"
link.add "style","margin-top:15px;margin-left:10px"
link.add "class","btn btn-info"
link.add "onclick","$('#mb_aspFormAction').val('link');$('#multiButtons').submit();return false"

dim button : set button=form.field
button.add "type","comment"
button.add "tag","button"
button.add "html","Button"
button.add "style","margin-top:15px;margin-left:10px"
button.add "class","btn btn-warning"
button.add "onclick","$('#mb_aspFormAction').val('button');$('#multiButtons').submit();return false;"

'##########################################################################################

form.build()

%>