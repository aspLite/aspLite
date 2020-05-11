<%
'load jquery date-functions (sets the dateformat!)
aspL("default/asp/includes/jQueryUiFunctions.asp")

'load includes (classes) for this particular example
aspL("default/asp/datatables/includes.asp")

'create an instance of the contact class
dim contact : set contact = new cls_contact
contact.pick(aspL.getRequest("iId"))

'create form
dim form : set form=aspl.form
form.initialize=false

'catch ajax call!
if form.postback then

	'session securitycheck
	if form.sameSession then
	
		dim reload : set reload=form.field("script")

		select case aspl.getRequest("btnDTTAction")
		
			case "save"
		
				contact.sText		= trim(left(aspl.getRequest("sText"),255))
				contact.iNumber		= aspl.getRequest("iNumber")			
				contact.dDate		= dateFromPicker(aspl.getRequest("dDate"))
				contact.iCountryID	= aspl.getRequest("iCountryID")
				
				if contact.save then
								
					set feedback=form.field("element")
					feedback.add "tag","div"
					feedback.add "html","Save OK"
					feedback.add "class","alert alert-success"				
					
					reload.add "text","$('#jsExampleSSP').DataTable().ajax.reload(null, false)"
							
				else
				
					set feedback=form.field("element")
					feedback.add "tag","div"			
					feedback.add "html","Error! Fill in all fields"
					feedback.add "class","alert alert-warning"
								
				end if			
			
			case "delete"
			
				contact.delete			
				
				reload.add "text","$('#jsExampleSSP').DataTable().ajax.reload(null, false);$('#exampleModal').modal('toggle')"
				
		
		end select
		
	end if
	
end if

dim iContactID : set iContactID = form.field("hidden")
iContactID.add "value",contact.iId
iContactID.add "name","iId"

dim sText : set sText = form.field("text")
sText.add "label","text"
sText.add "class","form-control"
sText.add "value",contact.sText
sText.add "required",true
sText.add "name","sText"

dim iNumber : set iNumber = form.field("number")
iNumber.add "label","number"
iNumber.add "class","form-control"
iNumber.add "value",contact.iNumber
iNumber.add "required",true
iNumber.add "name","iNumber"

dim countryList : set countryList=new cls_countryList

dim iCountryID : set iCountryID = form.field("select")
iCountryID.add "label","country"
iCountryID.add "class","form-control"
iCountryID.add "value",contact.iCountryID
iCountryID.add "required",true
iCountryID.add "name","iCountryID"
iCountryID.add "options",countryList.list

dim today : set today=form.field("text")
today.add "label","Select a jQuery UI date"
today.add "name","dDate"
today.add "class","form-control"
today.add "id","df1"
today.add "value",dateToPicker(contact.dDate)

dim script : set script=form.field("script")
script.add "text","$('#df1').datepicker({inline: true, dateFormat:'" & dateformat & "'})"

dim btnDTTAction : set btnDTTAction=form.field("hidden")
btnDTTAction.add "name","btnDTTAction"
btnDTTAction.add "id","btnDTTAction"

dim save : set save = form.field("button")
save.add "html","Save"
save.add "class","btn btn-primary"
save.add "style","margin-top:10px"
save.add "onclick","$('#btnDTTAction').val('save');aspAjax('POST',aspLiteAjaxHandler,$('#" & form.id & "').serialize(),aspForm)"

if aspl.convertNmbr(contact.iId)<>0 then
	dim delete : set delete = form.field("button")
	delete.add "html","Delete"	
	delete.add "class","btn btn-warning"
	delete.add "style","margin-top:10px;margin-left:5px"
	delete.add "container","span"
	delete.add "onclick","if (confirm('Are you sure?')) {$('#btnDTTAction').val('delete');aspAjax('POST',aspLiteAjaxHandler,$('#" & form.id & "').serialize(),aspForm)}"
end if

form.build

%>