<%@ OutputCache Duration="900" VaryByParam="*" %>
<%@ Page Debug="false" %>
<%@ Import Namespace="System.Drawing" %>
<%@ Import Namespace="System.Drawing.Imaging" %>
<%@ Import Namespace="System.IO" %>
<script language="VB" runat="server">
	
    Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
        
        On Error Resume Next		
             
        'retrieve relative path to image
        Dim imageUrl As String = HttpUtility.UrlDecode(Request.QueryString("img"))	
		dim doresize As Boolean
		dim isPNG As Boolean
		isPNG=false
		doresize=true
		
		if Lcase(Right(imageUrl,3))="png" then
			isPNG=true
		else
			isPNG=false
		end if             
		
		if isPNG then
			response.contenttype="image/png"
		else
			response.contenttype="image/jpeg"
		end if            
		
		'prepare thumbnail
		Dim fullSizeImg As System.Drawing.Image
		fullSizeImg = System.Drawing.Image.FromFile(Server.MapPath(imageUrl))						
            
        If fullSizeImg Is Nothing Then Response.End()            
        
		'??
		Dim dummyCallBack As System.Drawing.Image.GetThumbnailImageAbort
		dummyCallBack = New System.Drawing.Image.GetThumbnailImageAbort(AddressOf ThumbnailCallback)
		
		Dim thumbNailImg As System.Drawing.Image
		Dim newWidth As Integer
		Dim newHeight As Integer
		Dim maxSize As Double
		If IsNumeric(Request.QueryString("maxSize")) And Request.QueryString("maxSize") <> "" Then
			maxSize = Request.QueryString("maxSize")
		Else
			maxSize = 0
		End If	
		
		if maxSize > 2560 then
			maxSize=2560
		end if
			
		'calculate new widht/height, if any  
	   
		If fullSizeImg.Width > maxSize Or fullSizeImg.Height > maxSize Then			
		
			'for better quality
			fullSizeImg.RotateFlip(System.Drawing.RotateFlipType.Rotate90FlipX)
			fullSizeImg.RotateFlip(System.Drawing.RotateFlipType.Rotate90FlipX)
		
			If fullSizeImg.Width >= fullSizeImg.Height Then
				newWidth = maxSize
				newHeight = (fullSizeImg.Height / fullSizeImg.Width) * maxSize
			Else
				newWidth = (fullSizeImg.Width / fullSizeImg.Height) * maxSize
				newHeight = maxSize
			End If
		Else
			
			newWidth = fullSizeImg.Width
			newHeight = fullSizeImg.Height
			doresize = False
			
		End If
		
		dim square as string = Request.QueryString("square")
		IF square IS NOTHING ANDALSO square.length = 0 Then
			square="0"
		END IF
		
		If square = "1" Then
			If newWidth < newHeight Then
				thumbNailImg = fullSizeImg.GetThumbnailImage(newWidth * (newHeight / newWidth), newHeight * (newHeight / newWidth), dummyCallBack, IntPtr.Zero)
				thumbNailImg = CropImage(thumbNailImg, New Point(0, (newHeight - newWidth) / 2), New Point(newHeight, newHeight + (newHeight - newWidth) / 2))
			Else
				thumbNailImg = fullSizeImg.GetThumbnailImage(newWidth * (newWidth / newHeight), newHeight * (newWidth / newHeight), dummyCallBack, IntPtr.Zero)
				thumbNailImg = CropImage(thumbNailImg, New Point((newWidth - newHeight) / 2, 0), New Point(newWidth + ((newWidth - newHeight) / 2), newWidth))
		   End If				
		Else
			If doresize=True Then
				thumbNailImg = fullSizeImg.GetThumbnailImage(newWidth, newHeight, dummyCallBack, IntPtr.Zero)
			else				
				thumbNailImg = fullSizeImg					
			End If
		End If	  
		
		createOutput(thumbNailImg,isPNG)	 
		
		'Clean up / Dispose...
		thumbNailImg.Dispose()
	
		'Clean up / Dispose...
		fullSizeImg.Dispose()       
	   

		On Error GoTo 0
		
    End Sub

    Sub createOutput(ByRef ImageObj As System.Drawing.Image,isPNG as Boolean)
        
        On Error Resume Next	
		
		Dim specialEffect as String = HttpUtility.UrlDecode(Request.QueryString("SE"))
		
		select case specialEffect
			case "1"
				PureBW (ImageObj)
			case "2"
				GrayScale (ImageObj)
			case "3"
				PixelLoopConvert (ImageObj)
			
		end select	
		
		if isPNG then
			ImageObj.Save(Response.OutputStream, ImageFormat.png)
		else
			ImageObj.Save(Response.OutputStream, ImageFormat.Jpeg)
		end if
		
        On Error GoTo 0
        
    End Sub   
	
	Private Function CropImage(ByVal OriginalImage As Bitmap, ByVal TopLeft As Point, ByVal BottomRight As Point) As Bitmap
        Dim btmCropped As New Bitmap((BottomRight.Y - TopLeft.Y), (BottomRight.X - TopLeft.X))
        Dim grpOriginal As Graphics = Graphics.FromImage(btmCropped)
  
        grpOriginal.DrawImage(OriginalImage, New Rectangle(0, 0, btmCropped.Width, btmCropped.Height), _
            TopLeft.X, TopLeft.Y, btmCropped.Width, btmCropped.Height, GraphicsUnit.Pixel)
        grpOriginal.Dispose()
  
        Return btmCropped
    End Function
   
    
    Function ThumbnailCallback() As Boolean
        
        Return False
    
    End Function
	
	Public Function PureBW(ByVal image As System.Drawing.Bitmap, Optional ByVal Mode As BWMode = BWMode.By_Lightness, Optional ByVal tolerance As Single = 0) As System.Drawing.Bitmap
        Dim x As Integer
        Dim y As Integer
        If tolerance > 1 Or tolerance < -1 Then
            Throw New ArgumentOutOfRangeException
            Exit Function
        End If
        For x = 0 To image.Width - 1 Step 1
            For y = 0 To image.Height - 1 Step 1
                Dim clr As Color = image.GetPixel(x, y)
                If Mode = BWMode.By_RGB_Value Then
                    If (CInt(clr.R) + CInt(clr.G) + CInt(clr.B)) > 383 - (tolerance * 383) Then
                        image.SetPixel(x, y, Color.White)
                    Else
                        image.SetPixel(x, y, Color.Black)
                    End If
                Else
                    If clr.GetBrightness > 0.5 - (tolerance / 2) Then
                        image.SetPixel(x, y, Color.White)
                    Else
                        image.SetPixel(x, y, Color.Black)
                    End If
                End If
            Next
        Next
        Return image
    End Function
	
    Enum BWMode
        By_Lightness
        By_RGB_Value
    End Enum
	
	Public Function GrayScale (ByVal image As System.Drawing.Bitmap)
		
		Dim X As Integer
		Dim Y As Integer
		Dim clr As Integer

		For X = 0 To image.Width - 1
			For Y = 0 To image.Height - 1
				clr = (CInt(image.GetPixel(X, Y).R) + _
					   image.GetPixel(X, Y).G + _
					   image.GetPixel(X, Y).B) \ 3
				image.SetPixel(X, Y, Color.FromArgb(clr, clr, clr))
			Next Y
		Next X
		
		return image
		
	End Function
	
	Public Function PixelLoopConvert (ByVal Image As System.Drawing.Bitmap)    
   
		For i As Integer = 0 To Image.Width - 1
		  For j As Integer = 0 To Image.Height - 1
			Dim iRed As Integer = Image.GetPixel(i, j).R
			Dim iGreen As Integer = Image.GetPixel(i, j).G
			Dim iBlue As Integer = Image.GetPixel(i, j).B

			' Bereken de nieuwe RGB-waarde voor deze pixel:
			Dim iSepiaRed As Integer = Math.Min(Convert.ToInt32(iRed * 0.393 + iGreen * 0.769 + iBlue * 0.189), 255)
			Dim iSepiaGreen As Integer = Math.Min(Convert.ToInt32(iRed * 0.349 + iGreen * 0.686 + iBlue * 0.168), 255)
			Dim iSepiaBlue As Integer = Math.Min(Convert.ToInt32(iRed * 0.272 + iGreen * 0.534 + iBlue * 0.131), 255)

			' Wijzig de RGB-waarde van de huidige pixel van de oude naar de nieuwe:
			Image.SetPixel(i, j, Color.FromArgb(iSepiaRed, iSepiaGreen, iSepiaBlue))
		  Next
		Next
		
		Return Image

	End Function
    
</script>