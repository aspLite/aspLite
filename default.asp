<!-- #include file="asplite/asplite.asp"-->
<%
aspL.exec("code/default.handler.asp") 'event-handler

'body
html=replace (html,"[BODY]",body,1,-1,1)

'timer
html=replace (html,"[TIMER]","<!--code took " & aspL.printTimer & " ms to execute-->",1,-1,1)
 
response.write html

'destroy aspLite
set aspL=nothing
%>