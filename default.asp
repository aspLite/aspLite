<!-- #include file="aspfw/aspfw.asp"-->
<%
dim asp
set asp=new cls_aspfw

asp.ASP_executeGlobal("code/default.asp") 'codebehind file

'resx files are never served to browsers, so they are safer to use
dim html
html=asp.ASP_loadfile("html/default.resx") 'html file (template)

'head
html=replace (html,"[TITLETAG]",titletag,1,-1,1)
html=replace (html,"[HEADJS]",asp.js.flushHEAD,1,-1,1)

'body
html=replace (html,"[BODYJS]",asp.js.flushBODY,1,-1,1)
html=replace (html,"[BODY]",body,1,-1,1)

response.write html

response.write vbcrlf & "<!--code took " & asp.printTimer & "ms to execute-->"

set asp=nothing
%>