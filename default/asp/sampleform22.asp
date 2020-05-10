<%

'load jquery date-functions (sets the dateformat!)
aspL("default/asp/includes/jQueryUiFunctions.asp")

'load includes (classes) for this particular example
aspL("default/asp/datatables/includes.asp")

'create an instance of the contact class
dim contact : set contact = new cls_contact
contact.pick(aspL.getRequest("iId"))

'catch ajax call!
if aspl.convertBool(aspl.getRequest("postBack")) then

	'prepare JSON answer	
	dim jsonDict : set jsonDict=aspL.dict
	jsonDict.add "iId",contact.iId

	select case aspl.getRequest("btnAction")
	
		case "save"
	
			contact.sText		= trim(left(aspl.getRequest("sText"),255))
			contact.iNumber		= aspl.getRequest("iNumber")
			contact.bBoolean	= aspl.getRequest("bBoolean")
			contact.dDate		= dateFromPicker(aspl.getRequest("dDate"))
			contact.iCountryID	= aspl.getRequest("iCountryID")
			
			if contact.save then
				jsonDict.add "status","OK"				
			else
				jsonDict.add "status","ERR"				
			end if			
		
		case "delete"
		
			contact.delete
			jsonDict.add "status","DELETE"	
	
	end select 
	
	'add the contact-object to the JSON output		
	contact.reflectTo(jsonDict)
	
	'the Ajax page-execution stops here!
	aspl.json.dump(jsonDict)
	
end if

dim form : form=aspL.loadText("default/html/sampleform22.resx")

'set the dateformat for jQuery UI DatePicker
form=replace(form,"[dateformat]",dateformat,1,-1,1)

'hide or show the delete button
if aspl.convertNmbr(contact.iId)=0 then 'new record!
	form=replace(form,"[display]"," style=""display:none"" ",1,-1,1)
else
	form=replace(form,"[display]","",1,-1,1)
end if

dim booleanlist : set booleanlist=new cls_booleanlist
dim countryList : set countryList=new cls_countryList

'form replacements
form=replace(form,"[iId]",aspL.sanitize(contact.iId),1,-1,1)
form=replace(form,"[sText]",aspL.sanitize(contact.sText),1,-1,1)
form=replace(form,"[iNumber]",aspL.sanitize(contact.iNumber),1,-1,1)
form=replace(form,"[booleanlist]",booleanlist.showSelected("option",aspL.sanitize(contact.bBoolean)),1,-1,1)
form=replace(form,"[dDate]",aspL.sanitize(dateToPicker(contact.dDate)),1,-1,1)
form=replace(form,"[countryList]",countryList.showSelected("option",aspL.sanitize(contact.iCountryID)),1,-1,1)

set contact=nothing
set booleanlist=nothing

aspl.json.dump(form)
%>