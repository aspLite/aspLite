<%
'load jquery date-functions (sets the dateformat!)
aspl.exec("code/demo_asp/functions.asp")

dim testData : set testData = new cls_testData
testData.pick(aspL.getRequest("iId"))

'catch ajax call!
if aspl.convertBool(aspl.getRequest("postBack")) then

	select case aspl.getRequest("btnAction")
	
		case "save"
	
			testData.sText		= trim(left(aspl.getRequest("sText"),255))
			testData.iNumber	= trim(left(aspl.getRequest("iNumber"),255))
			testData.bBoolean	= left(aspl.getRequest("bBoolean"),255)
			testData.dDate		= left(dateFromPicker(aspl.getRequest("dDate")),255)
			
			if testData.save then
				aspL.dump "OK"
			else
				aspL.dump "ERR"
			end if
		
		case "delete"
		
			testData.delete			
			aspL.dump "DELETE"
	
	end select 
	
end if

dim template : template=aspL.loadText("html/demo_asp/postdtt.resx")
template=replace(template,"[dateformat]",dateformat,1,-1,1)

if aspl.convertNmbr(testData.iId)=0 then 'new record!
	template=replace(template,"[display]"," style='display:none' ",1,-1,1)
end if

dim booleanlist : set booleanlist=new cls_booleanlist

'replacements
template=replace(template,"[iId]",testData.iId,1,-1,1)
template=replace(template,"[sText]",testData.sText,1,-1,1)
template=replace(template,"[iNumber]",testData.iNumber,1,-1,1)
template=replace(template,"[booleanlist]",booleanlist.showSelected("option",testData.bBoolean),1,-1,1)
template=replace(template,"[dDate]",dateToPicker(testData.dDate),1,-1,1)

set booleanlist=nothing

aspL.dump template

class cls_testData

	private db, rs
	public iId, sText, iNumber, bBoolean, dDate
	
	Private Sub Class_Initialize
	
		set db=aspl.plugin("database") : db.path="db/sample.mdb"
		
	End Sub
	
	Private Sub Class_Terminate
	
		Set db = nothing
		
	End Sub

	public function pick(id)
	
		if aspL.convertNmbr(id)<>0 then	
			
			set rs=db.rs : rs.open ("select * from testbigdata where iId=" & id)
			
			if not rs.eof then
			
				iId			=	rs("iId")
				sText		=	rs("sText")
				iNumber		=	rs("iNumber")
				bBoolean	=	rs("bBoolean")
				dDate		=	rs("dDate")
			
			end if
					
			set rs=nothing
		
		end if
	
	end function
	
	private function check
	
		check=true
		
		if aspL.isEmpty(sText) then 		check=false : exit function
		if not aspL.isNumber(iNumber) then 	check=false : exit function
		if aspL.isEmpty(bBoolean) then 		check=false : exit function		
		if aspL.isEmpty(dDate) then 		check=false : exit function		
		
	
	end function
	
	public function save
	
		if check then
		
			set rs=db.rs 

			if aspl.isNumber(iId) then
				rs.Open "select * from testbigdata where iId="& aspl.convertNmbr(iId)		
			else
				rs.Open "select * from testbigdata where 1=2"
				rs.AddNew				
			end if		
			
			rs("sText")		= sText
			rs("iNumber")	= iNumber
			rs("bBoolean")	= aspl.convertBool(bBoolean)
			rs("dDate")		= dDate
			
			rs.update()
			
			iId = aspl.convertNmbr(rs("iId"))
			
			rs.close		
			
			set rs=nothing
			
			save=true		
		
		else
		
			save=false
		
		end if
	
	end function
	
	
	public function delete	
		
		db.execute("delete from testbigdata where iId=" & aspl.convertNmbr(iId))
		
	end function

end class

Class cls_booleanlist
	
	Public list
	
	Private Sub Class_Initialize
	
		Set list = aspL.dict
		
		list.Add true, "true"
		list.Add false, "false"	
		
	End Sub
	
	Private Sub Class_Terminate
	
		Set list = nothing
		
	End Sub
	
	Public Function showSelected(mode, selected)		
		
		Select Case mode
		
			Case "single" : showSelected = list(selected)
			
			Case "option"				
				
				Dim key
				For each key in list
				
					showSelected = showSelected & "<option value=""" & key & """"
					
					If aspL.convertBool(selected) = aspL.convertBool(key) Then
						showSelected = showSelected & " selected"
					End If
					
					showSelected = showSelected & ">" & list(key) & "</option>" & vbcrlf
				Next
				
		End Select
		
	End Function
	
End class

%>