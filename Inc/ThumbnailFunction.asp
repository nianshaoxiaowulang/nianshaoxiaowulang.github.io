<%
'Ϊ�ļ����ˮӡ
Function AddWaterMark(FileName)
	Dim strMarkSettingSql,MarkSettingRs,objFileSystem,strFileExtName,objImage
	If InStr(FileName,":") = 0 Then												'���ļ���ת��Ϊʵ��·��
		FileName = Server.Mappath(FileName)
	End if
	If FileName <> "" and not IsNull(FileName) Then								'�ļ����Ƿ�Ϊ��,�����˳�
		strFileExtName = ""
		If InStr(FileName,".") <> 0 Then
			strFileExtName = Lcase(Trim(Mid(FileName,InStrRev(FileName,".")+1)))
		End if
		If strFileExtName <> "jpg" and strFileExtName <> "gif" and strFileExtName <> "bmp" and strFileExtName <> "png" Then'�ļ����ǿ���ͼƬ���˳�
			Exit Function
		End if
		Set objFileSystem = Server.CreateObject("Scripting.FileSystemObject")
		If objFileSystem.FileExists(FileName) Then				'�ļ�����,�����˳�
			strMarkSettingSql = "select * from FS_config"
			Set MarkSettingRs = conn.Execute(strMarkSettingSql)
			If MarkSettingRs("MarkComponent") <> "0" Then						'ѡ����ĳ��ˮӡ���,�����˳�
				Select Case MarkSettingRs("MarkComponent")
					Case "1"													'ʹ��AspJpeg���												
						If IsObjInstalled("Persits.Jpeg") Then					'AspJpeg����Ѱ�װ,�����˳�
							If IsExpired("Persits.Jpeg") Then
								Response.Write("Persits.Jpeg����ѹ��ڣ���ѡ�����������ر�ˮӡ���ܡ�")
								Response.End
							End if
							If MarkSettingRs("MarkType") = "1" Then				'�������ˮӡ
								AddTextMark 1,MarkSettingRs("MarkText"),MarkSettingRs("MarkFontColor"),MarkSettingRs("MarkFontName"),MarkSettingRs("MarkFontBond"),MarkSettingRs("MarkFontSize"),MarkSettingRs("MarkPosition"),FileName
							Else												'���ͼƬˮӡ
								AddPictureMark 1,MarkSettingRs("MarkWidth"),MarkSettingRs("MarkHeight"),MarkSettingRs("MarkPicture"),MarkSettingRs("MarkOpacity"),MarkSettingRs("MarkTranspColor"),MarkSettingRs("MarkPosition"),FileName
							End if
						End if
					Case "2"													'ʹ��wsImage���
						If strFileExtName = "png" Then							'wsImage�����֧��PNG�ļ�,�����˳�
							Exit Function
						End if
						If IsObjInstalled("wsImage.Resize") Then				'wsImage����Ѱ�װ,�����˳�
							If IsExpired("wsImage.Resize") Then
								Response.Write("wsImage.Resize����ѹ��ڣ���ѡ�����������ر�ˮӡ���ܡ�")
								Response.End
							End if
							If MarkSettingRs("MarkType") = "1" Then				'�������ˮӡ
								AddTextMark 2,MarkSettingRs("MarkText"),MarkSettingRs("MarkFontColor"),MarkSettingRs("MarkFontName"),MarkSettingRs("MarkFontBond"),MarkSettingRs("MarkFontSize"),MarkSettingRs("MarkPosition"),FileName
							Else												'���ͼƬˮӡ
								AddPictureMark 2,MarkSettingRs("MarkWidth"),MarkSettingRs("MarkHeight"),MarkSettingRs("MarkPicture"),MarkSettingRs("MarkOpacity"),MarkSettingRs("MarkTranspColor"),MarkSettingRs("MarkPosition"),FileName
							End if
						End if
					Case "3"													'ʹ��SA-ImgWriter���
						If IsObjInstalled("SoftArtisans.ImageGen") Then			'SA-ImgWriter����Ѱ�װ,�����˳�
							If IsExpired("SoftArtisans.ImageGen") Then
								Response.Write("SoftArtisans.ImageGen����ѹ��ڣ���ѡ�����������ر�ˮӡ���ܡ�")
								Response.End
							End if
							If MarkSettingRs("MarkType") = "1" Then				'�������ˮӡ
								AddTextMark 3,MarkSettingRs("MarkText"),MarkSettingRs("MarkFontColor"),MarkSettingRs("MarkFontName"),MarkSettingRs("MarkFontBond"),MarkSettingRs("MarkFontSize"),MarkSettingRs("MarkPosition"),FileName
							Else												'���ͼƬˮӡ
								AddPictureMark 3,MarkSettingRs("MarkWidth"),MarkSettingRs("MarkHeight"),MarkSettingRs("MarkPicture"),MarkSettingRs("MarkOpacity"),MarkSettingRs("MarkTranspColor"),MarkSettingRs("MarkPosition"),FileName
							End if
						End if
				End Select
			End if
			Set MarkSettingRs = nothing
		End if
		Set objFileSystem = nothing
	End if
End Function
'ΪͼƬ�������ˮӡ
Function AddTextMark(MarkComponentID,MarkText,MarkFontColor,MarkFontName,MarkFontBond,MarkFontSize,MarkPosition,FileName)
	Dim objImage,X,Y,Text,TextWidth,FontColor,FontName,FondBond,FontSize,OriginalWidth,OriginalHeight
	If InStr(FileName,":") = 0 Then																'���ļ���ת��Ϊʵ��·��
		FileName = Server.Mappath(FileName)
	End if
	Text = Trim(MarkText)
	If Text = "" Then
		Exit Function
	End if
	FontColor = Replace(MarkFontColor,"#","&H")
	FontName = MarkFontName
	If MarkFontBond = "1" Then
		FondBond = True
	Else
		FondBond = False
	End if
	FontSize = Cint(MarkFontSize)
	
	Select Case MarkComponentID
		Case 1
			If Not IsObjInstalled("Persits.Jpeg") Then
				Exit Function
			End if
			Set objImage = Server.CreateObject("Persits.Jpeg")
			objImage.Open FileName
			objImage.Canvas.Font.Color = FontColor
			objImage.Canvas.Font.Family = FontName
			objImage.Canvas.Font.Bold = FondBond
			objImage.Canvas.Font.Size = FontSize
			TextWidth = objImage.Canvas.GetTextExtent(Text)										'����GB2313������ַ�����ռ���
			
			If objImage.OriginalWidth < TextWidth Or objImage.OriginalHeight < FontSize Then	'���ͼƬ�߶�С�������С����С���ַ���������˳�
				Exit Function
			End if
			GetPostion Cint(MarkPosition),X,Y,objImage.OriginalWidth,objImage.OriginalHeight,TextWidth,FontSize '��������
			objImage.Canvas.Print X, Y, Text,134
			objImage.Save FileName
		Case 2
			If Not IsObjInstalled("wsImage.Resize") Then
				Exit Function
			End if
			Set objImage = Server.CreateObject("wsImage.Resize")
			objImage.LoadSoucePic Cstr(FileName)
			objImage.TxtMarkFont = CStr(FontName)
			objImage.TxtMarkBond = FondBond
			objImage.TxtMarkHeight = FontSize
			'objImage.GetSourceInfo OriginalWidth,OriginalHeight
			'GetPostion Cint(MarkPosition),X,Y,OriginalWidth,OriginalHeight,Len(Text)*FontSize*3/4,FontSize '��������
			FontColor = "&H"&Mid(FontColor,7)&Mid(FontColor,5,2)&Mid(FontColor,3,2)				'��ɫ����ת��&HBBGGRR
			objImage.AddTxtMark Cstr(FileName),CStr(Text),Clng(FontColor),1,1
		Case 3
			If Not IsObjInstalled("SoftArtisans.ImageGen") Then
				Exit Function
			End if
			Set objImage = Server.CreateObject("SoftArtisans.ImageGen")
			objImage.LoadImage FileName
			objImage.Font.height = FontSize
			objImage.Font.name	= FontName
			FontColor = "&H"&Mid(FontColor,7)&Mid(FontColor,5,2)&Mid(FontColor,3,2)				'��ɫ����ת��&HBBGGRR
			objImage.Font.Color	= Clng(FontColor)
			objImage.Text = Text
			GetPostion Cint(MarkSettingRs("MarkPosition")),X,Y,objImage.Width,objImage.Height,objImage.TextWidth,objImage.TextHeight '��������
			objImage.DrawTextOnImage X, Y,objImage.TextWidth,objImage.TextHeight
			objImage.SaveImage 0, objImage.ImageFormat, FileName 
	End Select
	Set objImage = nothing
End Function
'ΪͼƬ���ͼƬˮӡ
Function AddPictureMark(MarkComponentID,MarkWidth,MarkHeight,MarkPicture,MarkOpacity,MarkTranspColor,MarkPosition,FileName)
	Dim objImage,objMark,X,Y,OriginalWidth,OriginalHeight,Position
	If InStr(FileName,":") = 0 Then																'���ļ���ת��Ϊʵ��·��
		FileName = Server.Mappath(FileName)
	End if
	If IsNull(MarkWidth) Or MarkWidth = "" Then
		MarkWidth = 0
	Else
		MarkWidth = Cint(MarkWidth)
	End if
	If IsNull(MarkHeight) Or MarkHeight = "" Then
		MarkHeight = 0
	Else
		MarkHeight = Cint(MarkHeight)
	End if
	If MarkPicture = "" Then
		Exit Function
	End if
	If IsNull(MarkOpacity) Or MarkOpacity = "" Then
		MarkOpacity = 1
	Else
		MarkOpacity = Csng(MarkOpacity)
	End if
	If MarkTranspColor <> "" Then																'ת����ɫ����
		MarkTranspColor = Replace(MarkTranspColor,"#","&H")
	Else
	End if
	Select Case MarkComponentID
		Case 1
			If Not IsObjInstalled("Persits.Jpeg") Then
				Exit Function
			End if
			Set objImage = Server.CreateObject("Persits.Jpeg")
			Set objMark = Server.CreateObject("Persits.Jpeg")
			objImage.Open FileName
			If objImage.OriginalWidth < MarkWidth Or objImage.OriginalHeight < MarkHeight Then	'���ͼƬ�߶�С��ˮӡ�߶Ȼ���С����ˮӡ������˳�
				Exit Function
			End if
			objMark.Open Server.Mappath(MarkPicture)
			GetPostion Cint(MarkPosition),X,Y,objImage.OriginalWidth,objImage.OriginalHeight,MarkWidth,MarkHeight '��������
			If MarkTranspColor <> "" Then
				objImage.DrawImage X,Y,objMark,MarkOpacity,MarkTranspColor
			else
				objImage.DrawImage X,Y,objMark,MarkOpacity
			End if
			objImage.Save FileName
		Case 2
			If Not IsObjInstalled("wsImage.Resize") Then
				Exit Function
			End if
			Set objImage = Server.CreateObject("wsImage.Resize")
			objImage.LoadSoucePic Cstr(FileName)
			objImage.LoadImgMarkPic Server.Mappath(MarkPicture)
			objImage.GetSourceInfo OriginalWidth,OriginalHeight
			GetPostion Cint(MarkPosition),X,Y,OriginalWidth,OriginalHeight,MarkWidth,MarkHeight '��������
			If MarkTranspColor = "" Then
				MarkTranspColor = 0
			Else
				MarkTranspColor = "&H"&Mid(MarkTranspColor,7)&Mid(MarkTranspColor,5,2)&Mid(MarkTranspColor,3,2)				'��ɫ����ת��&HBBGGRR
			End if
			objImage.AddImgMark Cstr(FileName),int(X),int(Y),Clng(MarkTranspColor),Int(CSng(MarkOpacity)*100)
		Case 3
			If Not IsObjInstalled("SoftArtisans.ImageGen") Then
				Exit Function
			End if
			Set objImage = Server.CreateObject("SoftArtisans.ImageGen")
			objImage.LoadImage FileName
			Select Case Cint(MarkSettingRs("MarkPosition"))
				Case 1
					Position = 3
				Case 2
					Position = 5
				Case 3
					Position = 1
				Case 4
					Position = 6
				Case 5
					Position = 8
			End Select
			If MarkTranspColor <> "" Then
				MarkTranspColor = "&H"&Mid(MarkTranspColor,7)&Mid(MarkTranspColor,5,2)&Mid(MarkTranspColor,3,2)				'��ɫ����ת��&HBBGGRR
				objImage.AddWatermark Server.MapPath(MarkPicture), Position,CSng(MarkOpacity),Clng(MarkTranspColor)
			else
				objImage.AddWatermark Server.MapPath(MarkPicture), Position,CSng(MarkOpacity)
			End if
			'Position:saiTopMiddle 0 saiCenterMiddle 1 saiBottomMiddle 2 saiTopLeft 3 saiCenterLeft 4 saiBottomLeft 5 saiTopRight 6 saiCenterRight 7 saiBottomRight 8 
			objImage.SaveImage 0, objImage.ImageFormat,FileName 
	End Select
	Set objImage = nothing
	Set objMark = nothing
End Function
'����ˮӡ���ͼƬ������
Function GetPostion(MarkPosition,X,Y,ImageWidth,ImageHeight,MarkWidth,MarkHeight)
	Select Case Cint(MarkPosition)
		Case 1
			X = 1
			Y = 1
		Case 2
			X = 1
			Y = Int(ImageHeight - MarkHeight - 1)
		Case 3
			X = Int((ImageWidth - MarkWidth)/2)
			Y = Int((ImageHeight - MarkHeight)/2)
		Case 4
			X = Int(ImageWidth - MarkWidth - 1)
			Y = 1
		Case 5
			X = Int(ImageWidth - MarkWidth - 1)
			Y = Int(ImageHeight - MarkHeight - 1)
	End Select						
End Function
'��ԭͼƬ���������ﱣ���������������ͼ
Function CreateThumbnailEx(FileName,ThumbnailFileName)
	Dim strSql,RsThumbnailSetting
	strSql = "Select ThumbnailComponent,RateTF,ThumbnailWidth,ThumbnailHeight,ThumbnailRate From FS_Config"
	Set RsThumbnailSetting = Conn.Execute(strSql)
	CreateThumbnailEx = False
	If RsThumbnailSetting("ThumbnailComponent") <> "0" and (not IsNull(RsThumbnailSetting("ThumbnailComponent")))Then
		If RsThumbnailSetting("RateTF") = "0" Then
			CreateThumbnailEx = CreateThumbnail(FileName,Cint(RsThumbnailSetting("ThumbnailWidth")),Cint(RsThumbnailSetting("ThumbnailHeight")),0,ThumbnailFileName)
		Else
			CreateThumbnailEx = CreateThumbnail(FileName,0,0,Csng(RsThumbnailSetting("ThumbnailRate")),ThumbnailFileName)
		End if
	End if
	Set RsThumbnailSetting = nothing
End Function
'��ԭͼƬ����ָ����Ⱥ͸߶ȵ�����ͼ
Function CreateThumbnail(FileName,Width,Height,Rate,ThumbnailFileName)
	Dim strSql,RsSetting,objImage,iWidth,iHeight,strFileExtName
	CreateThumbnail = False
	If IsNull(FileName) Then									'���ԭͼƬδָ��ֱ���˳�
		Exit Function
	Elseif FileName="" Then
		Exit Function
	End if
	If InStr(FileName,".") <> 0 Then
		strFileExtName = Lcase(Trim(Mid(FileName,InStrRev(FileName,".")+1)))
	End if
	If strFileExtName <> "jpg" and strFileExtName <> "gif" and strFileExtName <> "bmp" and strFileExtName <> "png" Then'�ļ����ǿ���ͼƬ���˳�
		Exit Function
	End if
	If IsNull(ThumbnailFileName) Then							'�������ͼδָ������·��ֱ���˳�
		Exit Function
	Elseif ThumbnailFileName="" Then
		Exit Function
	End if
	If IsNull(Width) Then										'�������ͼ���δָ������ָ��Ϊ0
		Width = 0
	Elseif Width="" Then
		Width = 0
	End if
	If IsNull(Rate) Then										'�������ͼ���ű���δָ������ָ��Ϊ0
		Rate = 0
	Elseif Rate="" Then
		Rate = 0
	End if
	If IsNull(Height) Then										'�������ͼ�߶�δָ������ָ��Ϊ0
		Height = 0
	Elseif Height="" Then
		Height = 0
	End if
	If InStr(FileName,":") = 0 Then								'ԭͼƬ·��ת��������·��
		FileName = Server.Mappath(FileName)
	End if
	If InStr(ThumbnailFileName,":") = 0 Then					'����ͼ·��ת��������·��
		ThumbnailFileName = Server.Mappath(ThumbnailFileName)
	End if
	Width = Cint(Width)
	Height = Cint(Height)
	Rate = CSng(Rate)
	
	strSql = "Select ThumbnailComponent From FS_Config"
	Set RsSetting = Conn.Execute(strSql)
	Select Case Cint(RsSetting("ThumbnailComponent"))
		Case 0													'����ͼ���ܹر�,�˳�
			Exit Function
		Case 1
			If Not IsObjInstalled("Persits.Jpeg") Then			'Persits.Jpegδ��װ,�˳�
				Exit Function
			End if
			If IsExpired("Persits.Jpeg") Then
				Response.Write("Persits.Jpeg����ѹ��ڣ���ѡ�����������ر�����ͼ���ܡ�")
				Response.End
			End if
			Set objImage = Server.CreateObject("Persits.Jpeg")
			objImage.Open FileName
			If Rate = 0 and (Width <> 0 Or Height<> 0) Then
				If Width < objImage.OriginalWidth And Height < objImage.OriginalHeight Then
					If Width = 0 and Height <> 0 Then
						objImage.Width = objImage.OriginalWidth/objImage.OriginalHeight*Height
						objImage.Height = Height
					Elseif Width <> 0 and Height = 0 Then
						objImage.Width = Width
						objImage.Height = objImage.OriginalHeight/objImage.OriginalWidth*Width
					ElseIf Width <> 0 and Height <> 0 Then
						objImage.Width = Width
						objImage.Height = Height
					End if
				End if
			Elseif  Rate <> 0 Then
				objImage.Width = objImage.OriginalWidth*Rate
				objImage.Height = objImage.OriginalHeight*Rate
			End if
			objImage.Save ThumbnailFileName
		Case 2
			If Not IsObjInstalled("wsImage.Resize") Then			'wsImage.Resizeδ��װ,�˳�
				Exit Function
			End if
			If IsExpired("wsImage.Resize") Then
				Response.Write("wsImage.Resize����ѹ��ڣ���ѡ�����������ر�����ͼ���ܡ�")
				Response.End
			End if
			If strFileExtName = "png" Then							'wsImage.Resize��֧��PNGͼƬ,�����˳�
				Exit Function
			End if
			Set objImage = Server.CreateObject("wsImage.Resize")
			objImage.LoadSoucePic CStr(FileName)
			If Rate = 0 and (Width <> 0 Or Height<> 0) Then
				objImage.GetSourceInfo iWidth,iHeight
				If Width < iWidth And Height < iHeight Then
					If Width = 0 and Height <> 0 Then
						objImage.OutputSpic CStr(ThumbnailFileName),0,Height,2
					Elseif Width <> 0 and Height = 0 Then
						objImage.OutputSpic CStr(ThumbnailFileName),Width,0,1
					ElseIf Width <> 0 and Height <> 0 Then
						objImage.OutputSpic CStr(ThumbnailFileName),Width,Height,0
					Else
						objImage.OutputSpic CStr(ThumbnailFileName),1,1,3
					End if
				Else
					objImage.OutputSpic CStr(ThumbnailFileName),1,1,3
				End if
			Elseif  Rate <> 0 Then
				objImage.OutputSpic CStr(ThumbnailFileName),Rate,Rate,3
			Else
				objImage.OutputSpic CStr(ThumbnailFileName),1,1,3
			End if
		Case 3
			If Not IsObjInstalled("SoftArtisans.ImageGen") Then		'SoftArtisans.ImageGenδ��װ,�˳�
				Exit Function
			End if
			If IsExpired("SoftArtisans.ImageGen") Then
				Response.Write("SoftArtisans.ImageGen����ѹ��ڣ���ѡ�����������ر�����ͼ���ܡ�")
				Response.End
			End if
			Set objImage = Server.CreateObject("SoftArtisans.ImageGen")
			objImage.LoadImage FileName
			If Rate = 0 and (Width <> 0 Or Height<> 0) Then
				If Width < objImage.Width And Height < objImage.Height Then
					If Width = 0 and Height <> 0 Then
						objImage.CreateThumbnail  ,Clng(Height),0,true
					Elseif Width <> 0 and Height = 0 Then
						objImage.CreateThumbnail  Clng(Width),objImage.Height/objImage.Width*Width,0,false
					ElseIf Width <> 0 and Height <> 0 Then
						objImage.CreateThumbnail  Clng(Width),Clng(Height),0,false
					End if
				End if
			Elseif  Rate <> 0 Then
				objImage.CreateThumbnail Clng(objImage.Width*Rate),Clng(objImage.Height*Rate),0,false
			End if
			objImage.SaveImage 0,objImage.ImageFormat,ThumbnailFileName
		Case 4
			If Not IsObjInstalled("CreatePreviewImage.cGvbox") Then		'CreatePreviewImage.cGvboxδ��װ,�˳�
				Exit Function
			End if
			set objImage = Server.CreateObject("CreatePreviewImage.cGvbox")
			objImage.SetImageFile = FileName							'imagenameԭʼ�ļ�������·��
			If Rate = 0 and (Width <> 0 Or Height<> 0) Then
				objImage.SetPreviewImageSize = Width					'Ԥ��ͼ���
			Elseif  Rate <> 0 Then
				objImage.SetPreviewImageSize = objImage.SetPreviewImageSize*Rate				'Ԥ��ͼ���
			End if
			objImage.SetSavePreviewImagePath = ThumbnailFileName		'Ԥ��ͼ���·��
			If objImage.DoImageProcess = False Then						'����Ԥ��ͼ���ļ�
				Exit Function
			End if
	End Select
	CreateThumbnail = True	
End Function
%>