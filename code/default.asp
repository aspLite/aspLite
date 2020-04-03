<%
'###########################################################################
'## include these 2 lines to start each codebehind file for security reasons
option explicit  
asp.codebehind() 
'###########################################################################

dim html,titletag,body

'load the template file
'resx files are never served to browsers, so they are safer to use
html=asp.load("html/default.resx")

titletag="aspLite"
body="body"
%>