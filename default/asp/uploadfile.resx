<%
on error resume next

dim uploadsDirVar, filesize, filename

'default path: uploads
uploadsDirVar = server.MapPath ("uploads") 

dim upload
Set upload = aspL.plugin("uploader")

upload.Save uploadsDirVar

dim fileKey, strMessage

for each fileKey in Upload.UploadedFiles.keys
	
	filesize=Upload.UploadedFiles(fileKey).Length
	filename=Upload.UploadedFiles(fileKey).filename
	
	'for security reasons, I immediately remove the uploaded files.
	Upload.UploadedFiles(fileKey).delete()
	
next

set upload=nothing

on error goto 0
%>