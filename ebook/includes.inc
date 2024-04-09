<!-- #include file="asplite/asplite.asp"-->
<%
dim asplEvent : asplEvent=aspl.getRequest("asplEvent")

if not aspl.isEmpty(asplEvent) then

	dim form : set form=aspl.form
	
	select case asplEvent
	
		case "myheader"
		%>
		<!-- #include file="asp/myheader.inc"-->
		<%
		case "myarticle"
		%>
		<!-- #include file="asp/myarticle.inc"-->
		<%
		case "myfooter"
		%>
		<!-- #include file="asp/myfooter.inc"-->
		<%	
	
	end select
	
else

	'no event, load the HTML, CSS and JavaScript
	response.write aspl.loadText("html/default.inc")
	
end if
%>