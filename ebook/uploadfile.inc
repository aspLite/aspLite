<%
dim uploadsDirVar : uploadsDirVar = server.MapPath ("/ebook/upload") 
dim upload  : set upload = aspL.plugin("uploader")
'upload.Save uploadsDirVar 'commented out!
set upload=nothing
%>