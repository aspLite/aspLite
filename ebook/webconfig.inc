<?XML version="1.0" encoding="UTF-8"?>

<configuration>

<system.webServer>
	
	<!-- set to <httpErrors> while developing and when facing errors -->
	<httpErrors errorMode="Custom">
		
		<remove statusCode="404" subStatusCode="-1" />
		
		<!-- change to the asp-file of your choice - most likely /default.asp-->
		<error statusCode="404" path="/default.asp" responseMode="ExecuteURL" />	
			
		
	</httpErrors>
	

</system.webServer>

</configuration>