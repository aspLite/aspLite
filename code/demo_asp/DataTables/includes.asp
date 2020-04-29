<%
class cls_contact

	private db, rs
	public iId, sText, iNumber, bBoolean, dDate, iCountryID
	public reflection
	
	Private Sub Class_Initialize
	
		set db=aspl.plugin("database") : db.path="db/sample.mdb"
		reflection = Array("iId","sText","iNumber","bBoolean","dDate","iCountryID")
		
	End Sub
	
	Private Sub Class_Terminate
	
		Set db = nothing
		
	End Sub

	public function pick(id)
	
		if aspL.convertNmbr(id)<>0 then	
			
			set rs=db.rs : rs.open ("select * from contact where iId=" & id)
			
			if not rs.eof then
			
				iId			=	rs("iId")
				sText		=	rs("sText")
				iNumber		=	rs("iNumber")
				bBoolean	=	rs("bBoolean")
				dDate		=	rs("dDate")
				iCountryID	=	rs("iCountryID")
			
			end if
					
			set rs=nothing
		
		end if
	
	end function
	
	private function check
	
		check=true
		
		if aspL.isEmpty(sText) then 			check=false : exit function
		if not aspL.isNumber(iNumber) then 		check=false : exit function
		if aspL.isEmpty(bBoolean) then 			check=false : exit function		
		if not isDate(dDate) then 				check=false : exit function
		if not aspL.isNumber(iCountryID) then 	check=false : exit function
		
	
	end function
	
	public function save
	
		if check then
		
			set rs=db.rs 

			if aspl.isNumber(iId) then
				rs.Open "select * from contact where iId="& aspl.convertNmbr(iId)		
			else
				rs.Open "select * from contact where 1=2"
				rs.AddNew				
			end if		
			
			rs("sText")			= sText
			rs("iNumber")		= aspL.convertNmbr(iNumber)
			rs("bBoolean")		= aspl.convertBool(bBoolean)
			rs("dDate")			= dDate
			rs("iCountryID")	= aspL.convertNmbr(iCountryID)
			
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
		
		db.execute("delete from contact where iId=" & aspl.convertNmbr(iId))
		
	end function
	
	
	public function reflectTo(byref dict)	
				
		for i=lbound(reflection) to ubound(reflection)
			dict.add "contact_" & reflection(i),eval(reflection(i))
		next
	
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


Class cls_countryList
	
	Public list
	private db,rs
	
	Private Sub Class_Initialize
	
		Set list = aspL.dict
		
		set db=aspl.plugin("database") : db.path="db/sample.mdb"
		set rs=db.execute("select * from country")
		
		while not rs.eof
			id=rs("iId")
			name=rs("sText")
			list.Add id, name
			rs.movenext
		wend
		
	End Sub
	
	Private Sub Class_Terminate
	
		Set list = nothing
		
	End Sub
	
	Public Function showSelected(mode, selected)		
		
		Select Case mode
		
			Case "single" : showSelected = list(selected)
			
			Case "option"	

				showSelected="<option></option>"
				
				Dim key
				For each key in list
				
					showSelected = showSelected & "<option value=""" & key & """"
					
					If aspL.convertNmbr(selected) = aspL.convertNmbr(key) Then
						showSelected = showSelected & " selected"
					End If
					
					showSelected = showSelected & ">" & list(key) & "</option>" & vbcrlf
				Next
				
		End Select
		
	End Function
	
End class
%>