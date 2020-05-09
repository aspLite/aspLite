<%
'this folder will be zipped by the browser!
dim folderToZip
folderToZip="uploads"

if not aspL.isEmpty(request.querystring("stream")) then

	'security check
	if left(request.querystring("stream"),len(folderToZip))=folderToZip then
		aspL.dump(aspl.loadText(request.querystring("stream")))
	else
		response.end 'not allowed
	end if
	
end if

'load the javascript file first
dim javascript
javascript=aspL.loadText("default/html/jszip.resx")	

'get filelist
dim filelist,filelistArr,objSuperFolder,filepath
set filelist=aspL.dict
set objSuperFolder = aspL.fso.getfolder(server.mappath(folderToZip))

ShowSubfolders (objSuperFolder)

set objSuperFolder=nothing

dim url
for each url in filelist

	if filelist(url)<>"blob" then
		filelistArr=filelistArr & vbtab & "urls.push(['" & aspL.sanitizeJS(url) & "',aspLiteAjaxHandler + '?asplEvent=sampleform16&stream=" & server.urlencode(url) &"'])" & vbcrlf
	else
		filelistArr=filelistArr & vbtab & "urls.push(['" & aspL.sanitizeJS(url) & "','" & aspL.sanitizeJS(url) & "'])" & vbcrlf
	end if
next

javascript=replace(javascript,"[filelist]",filelistArr,1,-1,1)

set filelist=nothing

dim form : set form=aspl.form
form.target="sampleform16"
form.onsubmit="return false"

dim scriptAsp : set scriptAsp=form.field("script")
scriptAsp.add "text",javascript

'##########################################################################################

dim download : set download=form.field("button")
download.add "id","blob"
download.add "html","click to download zip"
download.add "class","btn btn-primary"

'##########################################################################################

form.build()


Sub ShowSubFolders(fFolder)

	dim objFolder,colFiles,Subfolder

	Set objFolder = aspL.fso.GetFolder(fFolder.Path)
	Set colFiles = objFolder.Files	
	
	For Each objFile in colFiles		
		
		if (objFile.Attributes & 2) then 'file is not hidden
			
			filepath=objFile.path
			filepath=replace(filepath,server.mappath(folderToZip),folderToZip,1,-1,1)			
			filepath=replace(filepath,"\","/",1,-1,1)			
			
			select case lcase(aspL.getFileType(objFile.name))
			
				case "htm","html","js","css","svg","json","txt","csv","xml","mobirise","less","rtf","log","scss","xsl","xsd","yml"
					
					filelist.add filepath,"text"		
				
				case "asp","asa","php","mdb","aspx","gitattributes","sample"
					'do nothing
					
				case else
				
					filelist.add filepath,"blob"
					
			end select
			
		end if
	Next

	For Each Subfolder in fFolder.SubFolders
		ShowSubFolders(Subfolder)
	Next
	
	set objFolder=nothing
	set colFiles=nothing
	
End Sub


%>