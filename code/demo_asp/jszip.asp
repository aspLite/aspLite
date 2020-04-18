<%

'this folder will be zipped by the browser!
dim folderToZip
folderToZip="uploads"

if not aspL.isLeeg(request.querystring("stream")) then

	'security check
	if left(request.querystring("stream"),len(folderToZip))=folderToZip then
		aspL.dump(aspl.loadText(request.querystring("stream")))
	else
		response.end 'not allowed
	end if
	
end if

'load the template file first (including JavScript CDN)
body=aspL.loadText("html/demo_asp/jszip.resx")	

'get filelist
dim filelist,filelistArr,objSuperFolder,filepath
set filelist=server.createobject("scripting.dictionary")
set objSuperFolder = aspL.fso.getfolder(server.mappath(folderToZip))

ShowSubfolders (objSuperFolder)

set objSuperFolder=nothing

dim url
for each url in filelist

	if filelist(url)<>"blob" then
		filelistArr=filelistArr & vbtab & "urls.push(['" & aspL.sanitizeJS(url) & "','demo.asp?e=jszip&stream=" & server.urlencode(url) &"'])" & vbcrlf
	else
		filelistArr=filelistArr & vbtab & "urls.push(['" & aspL.sanitizeJS(url) & "','" & aspL.sanitizeJS(url) & "'])" & vbcrlf
	end if
next

body=replace(body,"[filelist]",filelistArr,1,-1,1)

set filelist=nothing

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