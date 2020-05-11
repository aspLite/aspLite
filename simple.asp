<!-- #include file="asplite/asplite.asp"-->

<%
dim form, field

'the code below is executed on the server
select case aspl.getRequest("asplEvent") 'data-asplEvent="clickLink"

	case "simpleDiv1"
	
		'this aspLite form binds to data-asplTarget="simpleDiv1"
		set form=aspl.form
		set field = form.field("plain")
		field.add "html","<p><i>Onload event!</i></p>"
		form.build()

	case "clickLink"
	
		'this aspLite form binds to data-asplTarget="simpleDiv2"
		set form=aspl.form
		
		set field = form.field("script")
		field.add "text","console.log('you can also send some JavaScript back to the browser')"	
		
		set field = form.field("element")
		field.add "tag","p"
		field.add "html","It's a start! Check the console (F12)"
		
		form.build()		
	
end select 

%>

<!doctype html>
<html lang="en">
  <head>
    <!-- Required meta tags -->
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
	
  </head>
  <body> 
  
  
	<h3>"Onload"-event triggers aspLite form</h3>
	<!-- class="asplForm" triggers a function in aspLite.js that calls this ASP page and replaces simpleDiv1 --> 	
	
	<div class="asplForm" id="simpleDiv1" style="color:Blue"></div>
		
	
	<h3>Trigger aspLite form</h3>
	<!-- class="ajaxlink" triggers a function in aspLite.js that calls this ASP page 
	and passes asplEvent (clickLink) and  asplTarget (simpleDiv2) -->
 	<a href="#" class="ajaxLink" data-asplTarget="simpleDiv2" data-asplEvent="clickLink">Click here</a>
	
	<div id="simpleDiv2" style="color:Red"></div>		
	
	
	<!-- jQuery  & aspLite JS -->
	<script src="default/js/jquery.js"></script>	
	<script>
		var aspLiteAjaxHandler='simple.asp'
	</script>
	<script src="asplite/asplite.js"></script>
	
  </body>
</html>