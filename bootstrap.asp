<!-- #include file="aspfw/ASPFW_CORE.asp"-->
<%
dim asp
set asp=new cls_aspfw

asp.ASP_executeGlobal("code/bootstrap.asp") 'codebehind file

'head
html=replace (html,"[TITLETAG]",titletag,1,-1,1)

'title
html=replace (html,"[PAGETITLE]",titletag,1,-1,1) 'i reuse the titlag here

'body
html=replace (html,"[BODY]",body,1,-1,1)

response.write html

response.write vbcrlf & "<!--code took " & asp.printTimer & "ms to execute-->"

set asp=nothing
%>