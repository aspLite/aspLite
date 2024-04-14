<%
public function recordset
		
	'returns a recordset
	set recordset=server.createobject("adodb.recordset")
	set recordset.ActiveConnection = Nothing
	recordset.CursorLocation=3
	recordset.CursorType=3
	recordset.LockType=4

end function
%>