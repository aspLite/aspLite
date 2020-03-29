<%
'###########################################################################
'## include these 2 lines to start each codebehind file for security reasons
option explicit  
asp.codebehind() 
'###########################################################################

dim html,titletag,body

'resx files are never served to browsers, so they are safer to use
html=asp.ASP_loadfile("html/bootstrap.resx")

'the title shows the result of a function of the plugin helloworld
titletag=asp.plugin("helloworld").hw()

dim i
for i=1 to 50
	'generate some random words with random lengths
	body=body & asp.randomtext(asp.randomnumber(10,20)) & " "
next		

%>