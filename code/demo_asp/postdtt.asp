<%
'load jquery date-functions (sets the dateformat!)
aspl.exec("code/demo_asp/functions.asp")

'load includes (classes) for this particular example
aspl.exec("code/demo_asp/datatables/includes.asp")

'create an instance of the testData class
dim testData : set testData = new cls_testData
testData.pick(aspL.getRequest("iId"))

'catch ajax call!
if aspl.convertBool(aspl.getRequest("postBack")) then

	'prepare JSON answer
	dim jsonObj : set jsonObj=aspL.plugin("json")
	dim jsonDict : set jsonDict=aspL.dict
	jsonDict.add "iId",testData.iId

	select case aspl.getRequest("btnAction")
	
		case "save"
	
			testData.sText		= trim(left(aspl.getRequest("sText"),255))
			testData.iNumber	= aspl.getRequest("iNumber")
			testData.bBoolean	= aspl.getRequest("bBoolean")
			testData.dDate		= dateFromPicker(aspl.getRequest("dDate"))
			
			if testData.save then
				jsonDict.add "status","OK"				
			else
				jsonDict.add "status","ERR"				
			end if			
		
		case "delete"
		
			testData.delete			
			jsonDict.add "status","DELETE"	
	
	end select 
	
	'add the testData-object to the JSON output		
	testData.reflectTo(jsonDict)
	
	'the Ajax page-execution stops here!
	aspL.dumpJson jsonObj.toJSON("aspL", jsonDict, false)
	
end if

dim form : form=aspL.loadText("html/demo_asp/postdtt.resx")

'set the dateformat for jQuery UI DatePicker
form=replace(form,"[dateformat]",dateformat,1,-1,1)

'hide or show the delete button
if aspl.convertNmbr(testData.iId)=0 then 'new record!
	form=replace(form,"[display]"," style=""display:none"" ",1,-1,1)
else
	form=replace(form,"[display]","",1,-1,1)
end if

dim booleanlist : set booleanlist=new cls_booleanlist

'form replacements
form=replace(form,"[iId]",aspL.sanitize(testData.iId),1,-1,1)
form=replace(form,"[sText]",aspL.sanitize(testData.sText),1,-1,1)
form=replace(form,"[iNumber]",aspL.sanitize(testData.iNumber),1,-1,1)
form=replace(form,"[booleanlist]",booleanlist.showSelected("option",aspL.sanitize(testData.bBoolean)),1,-1,1)
form=replace(form,"[dDate]",aspL.sanitize(dateToPicker(testData.dDate)),1,-1,1)

set testData=nothing
set booleanlist=nothing

aspL.dump form
%>