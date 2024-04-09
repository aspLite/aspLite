<!-- #include file="aspLite/aspLite.asp"-->
<%
dim html : html=aspl.loadText("html/default.inc")

html=replace(html,"[titletag]","My first aspLite page",1,-1,1)
html=replace(html,"[body]","Hello World",1,-1,1)

aspl.dump(html)
%>