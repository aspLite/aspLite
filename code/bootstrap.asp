<%
option explicit  'include this line in all your code files - it ensures that you declare all variables you'll use.
asp.codebehind() 'include this line in all your code files - it ensures that this page does not load when called directly.

dim html,titletag,body
'resx files are never served to browsers, so they are safer to use
html=asp.ASP_loadfile("html/bootstrap.resx")

titletag="Bootstrap Starter Page in ASP/VBScript framework"

dim i
body=""
for i=1 to 100
	'generate some random words with random lengths
	body=body & asp.randomtext(asp.randomnumber(5,20)) & " "
next		

%>