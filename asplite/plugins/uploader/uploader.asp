<%
'  For examples, documentation, and your own free copy, go to:
'  http://www.freeaspupload.net
'  Note: You can copy and use this script for free and you can make changes
'  to the code, but you cannot remove the above comment.

'Changes:
'Aug 2, 2005: Add support for checkboxes and other input elements with multiple values
'Jan 6, 2009: Lars added ASP_CHUNK_SIZE
'Sep 3, 2010: Enforce UTF-8 everywhere; new function to convert byte array to unicode string

Class cls_asplite_uploader

	Public UploadedFiles, FormElements, errorMessage, overWriteFiles
	Private VarArrayBinRequest, StreamRequest, uploadedYet, internalChunkSize

	Private Sub Class_Initialize()
		Set UploadedFiles = aspL.dict
		Set FormElements = aspL.dict
		Set StreamRequest = Server.CreateObject("ADODB.Stream")
		StreamRequest.Type = 2 'adTypeText
		StreamRequest.Open
		uploadedYet = false
		overWriteFiles=false
		internalChunkSize = 150000
	End Sub
	
	Private Sub Class_Terminate()
		If IsObject(UploadedFiles) Then
			UploadedFiles.RemoveAll()
			Set UploadedFiles = Nothing
		End If
		If IsObject(FormElements) Then
			FormElements.RemoveAll()
			Set FormElements = Nothing
		End If
		StreamRequest.Close
		Set StreamRequest = Nothing
	End Sub

	Public Property Get Form(sIndex)
		Form = ""
		If FormElements.Exists(LCase(sIndex)) Then Form = FormElements.Item(LCase(sIndex))
	End Property

	Public Property Get Files()
		Files = UploadedFiles.Items
	End Property
	
    Public Property Get Exists(sIndex)
            Exists = false
            If FormElements.Exists(LCase(sIndex)) Then Exists = true
    End Property
	
	Function GetFileName(strSaveToPath, FileName)
	
		on error resume next
		
		Dim Counter, flag, strTempFileName, FileExt, NewFullPath, p
	   
		Counter = 0
		p = instrrev(FileName, ".")
		FileExt = mid(FileName, p+1)
		strTempFileName = left(FileName, p-1)
		NewFullPath = strSaveToPath & "\" & FileName
		Flag = False	
			
		Do Until Flag = True
			If aspL.fso.FileExists(NewFullPath) = False Then
				Flag = True
				GetFileName = Mid(NewFullPath, InstrRev(NewFullPath, "\") + 1)				
			Else
				Counter = Counter + 1
				NewFullPath = strSaveToPath & "\" & strTempFileName & Counter & "." & FileExt
			End If
		Loop		 
		 
		on error goto 0		
		
	End Function 
            
	Public Sub Save(path)
	
		on error resume next
		
		Dim streamFile, fileItem, filePath

		if Right(path, 1) <> "\" then path = path & "\"

		if not uploadedYet then Upload

		For Each fileItem In UploadedFiles.Items		
		
			if overWriteFiles then
				filePath = path & fileItem.FileName
			else
				filePath = path & GetFileName(path,fileItem.FileName)
			end if
			
			Set streamFile = Server.CreateObject("ADODB.Stream")
			streamFile.Type = 1 'adTypeBinary
			streamFile.Open
			StreamRequest.Position=fileItem.Start
			StreamRequest.CopyTo streamFile, fileItem.Length
			streamFile.SaveToFile filePath, 2 'adSaveCreateOverWrite
			streamFile.close
			Set streamFile = Nothing
			fileItem.Path = filePath			
			
		 Next		
		 
		on error goto 0
		 
	End Sub	

	Public Sub Upload()
	
		Dim nCurPos, nDataBoundPos, nLastSepPos, nPosFile, nPosBound, sFieldName, osPathSep, auxStr, readBytes, readLoop, tmpBinRequest
		
		'RFC1867 Tokens
		Dim vDataSep
		Dim tNewLine, tDoubleQuotes, tTerm, tFilename, tName, tContentDisp, tContentType
		tNewLine = String2Byte(Chr(13))
		tDoubleQuotes = String2Byte(Chr(34))
		tTerm = String2Byte("--")
		tFilename = String2Byte("filename=""")
		tName = String2Byte("name=""")
		tContentDisp = String2Byte("Content-Disposition")
		tContentType = String2Byte("Content-Type:")

		uploadedYet = true

		on error resume next
			' Copy binary request to a byte array, on which functions like InstrB and others can be used to search for separation tokens
			readBytes = internalChunkSize
			VarArrayBinRequest = Request.BinaryRead(readBytes)
			VarArrayBinRequest = midb(VarArrayBinRequest, 1, lenb(VarArrayBinRequest))
			Do Until readBytes < 1
				tmpBinRequest = Request.BinaryRead(readBytes)
				if readBytes > 0 then
					VarArrayBinRequest = VarArrayBinRequest & midb(tmpBinRequest, 1, lenb(tmpBinRequest))
				end if
			Loop
			StreamRequest.WriteText(VarArrayBinRequest)
			StreamRequest.Flush()
			if Err.Number <> 0 then 
				errorMessage="<p>System reported an error:</p>"
				errorMessage=errorMessage & "<p>" & Err.Description  & "</p>"
				errorMessage=errorMessage & "<p>The most likely cause for this error is the incorrect setup of AspMaxRequestEntityAllowed in IIS MetaBase.</p>"
				Exit Sub
			end if
		on error goto 0 'reset error handling

		nCurPos = FindToken(tNewLine,1) 'Note: nCurPos is 1-based (and so is InstrB, MidB, etc)

		If nCurPos <= 1  Then Exit Sub
		 
		'vDataSep is a separator like -----------------------------21763138716045
		vDataSep = MidB(VarArrayBinRequest, 1, nCurPos-1)

		'Start of current separator
		nDataBoundPos = 1

		'Beginning of last line
		nLastSepPos = FindToken(vDataSep & tTerm, 1)

		Do Until nDataBoundPos = nLastSepPos
			
			nCurPos = SkipToken(tContentDisp, nDataBoundPos)
			nCurPos = SkipToken(tName, nCurPos)
			sFieldName = ExtractField(tDoubleQuotes, nCurPos)

			nPosFile = FindToken(tFilename, nCurPos)
			nPosBound = FindToken(vDataSep, nCurPos)
			
			If nPosFile <> 0 And  nPosFile < nPosBound Then
				Dim oUploadFile
				Set oUploadFile = New UploadedFile
				
				nCurPos = SkipToken(tFilename, nCurPos)
				auxStr = ExtractField(tDoubleQuotes, nCurPos)
                ' We are interested only in the name of the file, not the whole path
                ' Path separator is \ in windows, / in UNIX
                ' While IE seems to put the whole pathname in the stream, Mozilla seem to 
                ' only put the actual file name, so UNIX paths may be rare. But not impossible.
                osPathSep = "\"
                if InStr(auxStr, osPathSep) = 0 then osPathSep = "/"
				oUploadFile.FileName = Right(auxStr, Len(auxStr)-InStrRev(auxStr, osPathSep))

				if (Len(oUploadFile.FileName) > 0) then 'File field not left empty
					nCurPos = SkipToken(tContentType, nCurPos)
					
                    auxStr = ExtractField(tNewLine, nCurPos)
                    ' NN on UNIX puts things like this in the stream:
                    '    ?? python py type=?? python application/x-python
					oUploadFile.ContentType = Right(auxStr, Len(auxStr)-InStrRev(auxStr, " "))
					nCurPos = FindToken(tNewLine, nCurPos) + 4 'skip empty line
					
					oUploadFile.Start = nCurPos+1
					oUploadFile.Length = FindToken(vDataSep, nCurPos) - 2 - nCurPos
					
					If oUploadFile.Length > 0 Then 
					
						'for security reasons, only a given number of filetypes can be uploaded
					
						select case lcase(oUploadFile.FileType)
							case "jpg","jpeg","jpe","jp2","jfif","gif","bmp","png","psd","eps","ico","tif","tiff","ai","raw","tga","mng","svg","doc","rtf","txt","wpd","wps","csv","xml","xsd","sql","pdf","xls","mdb","ppt","docx","xlsx","pptx","ppsx","artx","mp3","wma","mid","midi","mp4","mpg","mpeg","wav","ram","ra","avi","mov","flv","m4a","m4v","htm","html","css","swf","js","rar","zip","ogv","ogg","webm","tar","gz","eot","ttf","ics","woff","cod","msg","odt"
								UploadedFiles.Add LCase(nCurPos & sFieldName), oUploadFile
							case else
								'not allowed
						end select						
					
					end if
				End If
			Else
				Dim nEndOfData, fieldValueUniStr
				nCurPos = FindToken(tNewLine, nCurPos) + 4 'skip empty line
				nEndOfData = FindToken(vDataSep, nCurPos) - 2
				fieldValueuniStr = ConvertUtf8BytesToString(nCurPos, nEndOfData-nCurPos)
				If Not FormElements.Exists(LCase(sFieldName)) Then 
					FormElements.Add LCase(nCurPos & sFieldName), fieldValueuniStr
				else
                    FormElements.Item(LCase(sFieldName))= FormElements.Item(LCase(sFieldName)) & ", " & fieldValueuniStr
                end if 

			End If

			'Advance to next separator
			nDataBoundPos = FindToken(vDataSep, nCurPos)
		Loop
	End Sub

	Private Function SkipToken(sToken, nStart)
		SkipToken = InstrB(nStart, VarArrayBinRequest, sToken)
		If SkipToken = 0 then		
			aspL.asperror("Error in parsing uploaded binary request. The most likely cause for this error is the incorrect setup of AspMaxRequestEntityAllowed in IIS MetaBase.")
		end if
		SkipToken = SkipToken + LenB(sToken)
	End Function

	Private Function FindToken(sToken, nStart)
		FindToken = InstrB(nStart, VarArrayBinRequest, sToken)
	End Function

	Private Function ExtractField(sToken, nStart)
		Dim nEnd
		nEnd = InstrB(nStart, VarArrayBinRequest, sToken)
		If nEnd = 0 then
			aspL.asperror("Error in parsing uploaded binary request.")		
		end if
		ExtractField = ConvertUtf8BytesToString(nStart, nEnd-nStart)
	End Function

	'String to byte string conversion
	Private Function String2Byte(sString)
		Dim i
		For i = 1 to Len(sString)
		   String2Byte = String2Byte & ChrB(AscB(Mid(sString,i,1)))
		Next
	End Function

	Private Function ConvertUtf8BytesToString(start, length)	
		StreamRequest.Position = 0
	
	    Dim objStream
	    Dim strTmp
	    
	    ' init stream
	    Set objStream = Server.CreateObject("ADODB.Stream")
	    objStream.Charset = "utf-8"
	    objStream.Mode = 3 'adModeReadWrite
	    objStream.Type = 1 'adTypeBinary
	    objStream.Open
	    
	    ' write bytes into stream
	    StreamRequest.Position = start+1
	    StreamRequest.CopyTo objStream, length
	    objStream.Flush
	    
	    ' rewind stream and read text
	    objStream.Position = 0
	    objStream.Type = 2 'adTypeText
	    strTmp = objStream.ReadText
	    
	    ' close up and return
	    objStream.Close
	    Set objStream = Nothing
	    ConvertUtf8BytesToString = strTmp	
	End Function
End Class

Class UploadedFile

	Public ContentType, Start, Length, Path
	Private nameOfFile

    ' Need to remove characters that are valid in UNIX, but not in Windows
    Public Property Let FileName(fN)
        nameOfFile = fN
        nameOfFile = SubstNoReg(nameOfFile, "\", "_")
        nameOfFile = SubstNoReg(nameOfFile, "/", "_")
        nameOfFile = SubstNoReg(nameOfFile, ":", "_")
        nameOfFile = SubstNoReg(nameOfFile, "*", "_")
        nameOfFile = SubstNoReg(nameOfFile, "?", "_")
        nameOfFile = SubstNoReg(nameOfFile, """", "_")
        nameOfFile = SubstNoReg(nameOfFile, "<", "_")
        nameOfFile = SubstNoReg(nameOfFile, ">", "_")
        nameOfFile = SubstNoReg(nameOfFile, "|", "_")
    End Property

    Public Property Get FileName()
        FileName = nameOfFile
    End Property
	
	Public Property Get FileType
	FileType=aspL.getFileType(FileName)	
	end property
	
	Public function delete()	
	
		if aspL.fso.FileExists (path) then
			aspL.fso.DeleteFile (path)
		end if	
    
   	end function
	
	
	Public function rename(newname)	
	
		if aspL.fso.FileExists (path) then
	
			if aspl.convertStr(newname)=aspl.convertStr(path) then exit function
			
			if aspL.fso.FileExists (newname) then aspL.fso.deletefile(newname)
						
			aspL.fso.MoveFile path, newname
			
		end if	
    
    end function
	
	' Does not depend on RegEx, which is not available on older VBScript
	' Is not recursive, which means it will not run out of stack space
	Function SubstNoReg(initialStr, oldStr, newStr)
		Dim currentPos, oldStrPos, skip
		If IsNull(initialStr) Or Len(initialStr) = 0 Then
			SubstNoReg = ""
		ElseIf IsNull(oldStr) Or Len(oldStr) = 0 Then
			SubstNoReg = initialStr
		Else
			If IsNull(newStr) Then newStr = ""
			currentPos = 1
			oldStrPos = 0
			SubstNoReg = ""
			skip = Len(oldStr)
			Do While currentPos <= Len(initialStr)
				oldStrPos = InStr(currentPos, initialStr, oldStr)
				If oldStrPos = 0 Then
					SubstNoReg = SubstNoReg & Mid(initialStr, currentPos, Len(initialStr) - currentPos + 1)
					currentPos = Len(initialStr) + 1
				Else
					SubstNoReg = SubstNoReg & Mid(initialStr, currentPos, oldStrPos - currentPos) & newStr
					currentPos = oldStrPos + skip
				End If
			Loop
		End If
	End Function	
    
End Class 
%>
