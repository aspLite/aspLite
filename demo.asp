<!-- #include file="asplite/asplite.asp"-->
<%
aspL.exec("code/demo.handler.asp") 'event-handler

'head
html=replace (html,"[TITLETAG]",titletag,1,-1,1)

'title
html=replace (html,"[PAGETITLE]",titletag,1,-1,1)

'body
html=replace (html,"[BODY]",body,1,-1,1)

'timer
html=replace (html,"[TIMER]","<!--code took " & aspL.printTimer & " ms to execute-->",1,-1,1)
 
response.write html

'destroy aspLite
set aspL=nothing
%>