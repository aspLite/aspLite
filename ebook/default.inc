<!-- #include file="asplite/asplite.asp"-->
<%
dim asplEvent : asplEvent=aspl.getRequest("asplEvent")

if not aspl.isEmpty(asplEvent) then
	
	'dynamically execute the scriptname in asplEvent
	aspl("asp/" & asplEvent & ".inc")
	
else

	'no event, load the HTML, CSS and JavaScript
	response.write aspl.loadText("html/default.inc")
	
end if
%>