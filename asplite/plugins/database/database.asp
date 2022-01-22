<%
Class cls_asplite_database

	public path,dbms
	public sqlserver,initial_catalog 'server instance name, database name
	public userID,password 'only for SQL server
	
	'dbms=1 -> Access (default)
	'dbms=2 -> SQL Server Integrated Security = SSPI (NO SQL Server username/password - sqlserver,initial_catalog REQUIRED)
	'dbms=3 -> SQL Server with userID & password  & sqlserver & initial_catalog ALL REQUIRED

	private p_getConn
	
	Private Sub Class_Initialize()
		dbms=1 '-> Access (default)
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
	

	Public Function rs
	
		On Error Resume Next
		
		Set rs = server.CreateObject ("adodb.recordset")
		rs.CursorType = 1
		rs.LockType = 3
		set rs.ActiveConnection = getConn()		
		
		On Error Goto 0
		
	End Function
	
	Private function getConn()
	
		'this is the crucial part of this class.
		'i always use the native OLEDB driver for Access & SQL Server (much MUCH faster than ODBC)
		'however, to get this working for Access, you need to enable 32 applications for the application pool in IIS
		'luckily, this is taken care of by most ISP's and it's the default behaviour in IIS Express
		'note that I create a connection object only ONCE through the lifespan of this ASP page
		'it's important to NOT open multiple connections to an Access database, as there is a limit
		'of 255 concurrent users. in the context of a web application however, where users are 
		'connected for a very short period of time (milliseconds), you will most likely never run 
		'into problems when using Access, even for very busy websites.
	
		On Error Resume Next		
		
		if p_getConn is nothing then
			
			Set p_getConn = Server.Createobject("ADODB.Connection")		
		
			select case aspL.convertNmbr(dbms)			
			
				'Access
				case 1 : p_getConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & server.mappath(path)
								
				'SQL Server SSPI - no username/password required
				case 2 : p_getConn.Open "Provider=SQLOLEDB;SERVER=" & sqlserver & ";initial catalog=" & initial_catalog  & ";Integrated Security=SSPI;"
				
				'SQL Server username/password required
				case 3 : p_getConn.Open "Provider=SQLOLEDB;SERVER=" & sqlserver & ";initial catalog=" & initial_catalog  & ";User ID=" & userID & ";Password=" & password
			
			end select 
			
		end if		

		if err.number<>0 then	
		
			dim errM
			errM="Something went wrong while connection to the database:<ul>"
			errM=errM &"<li>You may need to enable 32 Bits application for your IIS application pool.</li>"
			errM=errM &"<li>The path to your database is not correct</li>"
			errM=errM &"<li>The username/password for SQL Server are incorrect</li>"
			errM=errM &"<li>IUSR has not sufficient permissions on the database file/folder</li>"
			errM=errM &"</ul><p>To be sure, check the error message below:</p>"
			
			aspL.asperror (errM)			
			
		end if

		set getConn=p_getConn

		On Error Goto 0
		
	end function
	
End Class
%>