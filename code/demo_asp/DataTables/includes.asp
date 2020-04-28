<%
class cls_testData

	private db, rs
	public iId, sText, iNumber, bBoolean, dDate
	public reflection
	
	Private Sub Class_Initialize
	
		set db=aspl.plugin("database") : db.path="db/sample.mdb"
		reflection = Array("iId","sText","iNumber","bBoolean","dDate")
		
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
	
	
	public function reflectTo(byref dict)	
				
		for i=lbound(reflection) to ubound(reflection)
			dict.add "testData_" & reflection(i),eval(reflection(i))
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
%>