<%
Class cls_asplite_accessDB

	public path

	private p_getConn
	
	Private Sub Class_Initialize()
		set p_getConn = nothing
	End Sub
	
	Private Sub Class_Terminate()
		set p_getConn = nothing
	End Sub
	
	Public Function Execute(sql)
	
		On Error Resume Next	
		
		dim connection : set connection = getConn()
		Set Execute = connection.Execute(sql)
		
		aspL.asperror "The query """ & server.htmlEncode(sql) & """ cannot be executed. There may be a connection error and/or a mistake in the SQL-syntax."
				
		On Error Goto 0
		
	End Function
	

	Public Function GetDynamicRS
	
		On Error Resume Next
		
		Set GetDynamicRS = server.CreateObject ("adodb.recordset")
		GetDynamicRS.CursorType = 1
		GetDynamicRS.LockType = 3
		set GetDynamicRS.ActiveConnection = getConn()		
		
		On Error Goto 0
		
	End Function
	
	Private function getConn()
	
		'this is the crucial part of this class.
		'i always use the native OLEDB driver for Access (much MUCH faster than ODBC)
		'however, to get this working, you need to enable 32 applications for the application pool in IIS
		'luckily, this is taken care of by most ISP's and it's the default behaviour in IIS Express
		'note that I create a connection object only ONCE through the lifespan of this ASP page
		'it's important to NOT open multiple connections to an Access database, as there is a limit
		'of 255 concurrent users. in the context of a web application however, where users are 
		'connected for a very short period of time (milliseconds), you will most likely never run 
		'into problems when using Access, even for very busy websites.
	
		On Error Resume Next
		
		if p_getConn is nothing then
			Set p_getConn = Server.Createobject("ADODB.Connection")		
			p_getConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & server.mappath(path)			
		end if			
		
		if err.number<>0 then	
		
			dim errM
			errM="Something went wrong while connection to the database:<ul>"
			errM=errM &"<li>You may need to enable 32 Bits application for your IIS application pool.</li>"
			errM=errM &"<li>The path to your database is not correct</li>"
			errM=errM &"<li>IUSR has not sufficient permissions on the database file/folder</li>"
			errM=errM &"</ul><p>To be sure, check the error message below:</p>"
			
			aspL.asperror (errM)			
			
		end if

		set getConn=p_getConn

		On Error Goto 0
		
	end function
	
End Class
%>