<!-- #include file="aspfw/aspfw.asp"-->
<%
dim aspfw
set aspfw=new cls_aspfw

aspfw.ASP_executeGlobal("code/default.resx") 'codebehind file

dim html
html=aspfw.ASP_loadfile("html/default.resx") 'html file (template)

'head
html=replace (html,"[TITLETAG]",titletag,1,-1,1)
html=replace (html,"[HEADJS]",javascript.flushHEAD,1,-1,1)

'body
html=replace (html,"[BODYJS]",javascript.flushBODY,1,-1,1)
html=replace (html,"[BODY]",body,1,-1,1)

response.write html

response.write timerObj.printTimer 'adds the number of milliseconds to the source

set aspfw=nothing
%>