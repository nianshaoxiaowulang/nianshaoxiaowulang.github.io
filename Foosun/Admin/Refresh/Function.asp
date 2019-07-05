<%
Function ReplaceAllServerFlag(Content)
	Dim regEx,Matches,Match,TempStr,ReturnValue,LoopVar
	Dim NotReplaceLable,ReplaceLableTF
	NotReplaceLable = Array("DownLoadList","ClassNewsList","LastClassPic","SpecialLastNewsList","Mall_LastClass","Mall_LastSpecialClass")
	Content = GetVisionStr & Content
	Set regEx = New RegExp
	regEx.Pattern = "{%=[^{%=}]*%}"
	regEx.IgnoreCase = True
	regEx.Global = True
	Set Matches = regEx.Execute(Content)
	ReplaceAllServerFlag = Content
	For Each Match in Matches
		TempStr = Match.Value
		TempStr = Replace(TempStr,Chr(13) & Chr(10),"")
		TempStr = Replace(TempStr,"{%=","")
		TempStr = Replace(TempStr,"%}","")
		TempStr = Left(TempStr,Instr(TempStr,"(")-1) & "," & Mid(TempStr,InStr(TempStr,"(")+1)
		TempStr = Left(TempStr,InStrRev(TempStr,")")-1)
		TempStr = Replace(TempStr,"""","")
		ReplaceLableTF = True
		for LoopVar = LBound(NotReplaceLable) to UBound(NotReplaceLable)
			if InStr(TempStr,NotReplaceLable(LoopVar)) <> 0 then 
				ReplaceLableTF = False
				if NotReplaceLableArray = "" then
					NotReplaceLableArray = TempStr
				else
					NotReplaceLableArray = NotReplaceLableArray & "$$$" & TempStr
				end if
				if NotReplaceLableOldArray = "" then
					NotReplaceLableOldArray = Match.Value
				else
					NotReplaceLableOldArray = NotReplaceLableOldArray & "$$$" & Match.Value
				end if
				Exit For
			end if
		Next
		if ReplaceLableTF = True then
			ReturnValue = GetLableContent(TempStr)
			ReplaceAllServerFlag = Replace(ReplaceAllServerFlag,Match.Value,ReturnValue)
		end if
	Next
End Function
Function ReplaceAllLable(Content)
'**************
'Replace Lable
'**************
Content=ReplaceYpren(Content)


	Dim RsLableObj,LableSql
	LableSql = "Select LableName,LableContent from FS_Lable"
	Set RsLableObj = Conn.Execute(LableSql)
	do while Not RsLableObj.Eof
		Content = Replace(Content,RsLableObj("LableName"),RsLableObj("LableContent"))
		RsLableObj.MoveNext
	Loop
	SEt RsLableObj = Nothing
	ReplaceAllLable = Content
End Function
Sub SaveFile(Content,LocalFileName)
	Dim AdodbStreamObj
	Set AdodbStreamObj = Server.CreateObject("Adodb.Stream")
	With AdodbStreamObj
		.Type = 2
		.Open
		.Charset = "GB2312"
		.WriteText replace(Content,WebDomain,"")
		.SaveToFile Server.MapPath(LocalFileName),2
		.Cancel()
		.Close()
	End With
	Set AdodbStreamObj = Nothing
End Sub
Sub FSOSaveFile(Content,LocalFileName)
	Dim FileObj,FilePionter
	'response.write LocalFileName
	'response.end
	Set FileObj=Server.CreateObject(G_FS_FSO)
	Set FilePionter = FileObj.CreateTextFile(Server.MapPath(LocalFileName),True) '创建文件
	FilePionter.Write Replace(Content,Webdomain,"")
	FilePionter.close     '释放对象
	Set FilePionter = Nothing
	Set FileObj = Nothing
End Sub
Function RefreshNews(NewsRecordSetObj) '生成新闻
	Dim SaveFilePath,FileContent,TempletFileName,RsClassObj,SaveNewsContent,FSOObj,FileStreamObj,FileObj
	Dim ContentArray,i,CurrPageNum,FileName,NewsContent,NewsPageStr,j,NewsPageCount,ClassSaveFilePath,TempFileContent
	Dim TempSysRootDir
	GetAvailableDoMain
	'/l
	if SysRootDir = "" then
		TempSysRootDir = ""
	else
		TempSysRootDir = "/" & SysRootDir
	end if
	'=====================================
	'添加上一篇，下一篇功能 modi 2005-07-15
	dim FS_NextTempStr,FS_PreviousTempStr
	Dim NowNewsID,NowClassID
	NowNewsID = NewsRecordSetObj("ID")
	NowClassID = NewsRecordSetObj("ClassID")
	'======================================
	SetRefreshValue "News",NewsRecordSetObj("NewsID")
	Set RsClassObj = Conn.Execute("Select DoMain,ClassEName,SaveFilePath from FS_NewsClass where ClassID='" & NewsRecordSetObj("ClassID") & "'")
	if Not RsClassObj.Eof then
		if RsClassObj("SaveFilePath") = "/" then
			ClassSaveFilePath = TempSysRootDir & RsClassObj("SaveFilePath")
		else
			ClassSaveFilePath = TempSysRootDir & RsClassObj("SaveFilePath") & "/"
		end if
		if Not IsNull(NewsRecordSetObj("NewsTemplet")) then
			TempletFileName = Server.MapPath(TempSysRootDir & NewsRecordSetObj("NewsTemplet"))
		else
			TempletFileName = ""
		end if
		Set FSOObj = Server.CreateObject(G_FS_FSO)
		if FSOObj.FileExists(TempletFileName) = False then
			FileContent = "模板不存在，请添加模板后再生成！"
		else
			Set FileObj = FSOObj.GetFile(TempletFileName)
			Set FileStreamObj = FileObj.OpenAsTextStream(1)
			if Not FileStreamObj.AtEndOfStream then
				FileContent = FileStreamObj.ReadAll
				if (NewsRecordSetObj("BrowPop") <> 0) And (NewsRecordSetObj("FileExtName") = "asp") then
					If Application("UseDatePath")="1" then 
						FileContent = GetPopStr(NewsRecordSetObj("BrowPop"),NewsRecordSetObj("path")&RsClassObj("SaveFilePath")) & FileContent
					else
						FileContent = GetPopStr(NewsRecordSetObj("BrowPop"),RsClassObj("SaveFilePath")) & FileContent
					end if
				end if
				FileContent = ReplaceAllLable(FileContent)
				'====================================
				'添加上一篇，下一篇功能 modi 2005-07-15
				Dim NextSql,NextRs,NextClassRs
				If Instr(FileContent,"{%=PrePageNews()%}")<>0 then 
					'上一篇
					NextSql = "Select TOP 1 FS_news.title,FS_news.Path,FS_news.FileName,FS_news.FileExtName,FS_NewsClass.ClassEName From FS_News,FS_newsclass where FS_News.HeadNewsTF=0 and FS_News.DelTF=0 and FS_News.ID < " & NowNewsID & " and FS_News.ClassID = '" & NowClassID & "' and FS_News.ClassID=FS_NewsClass.ClassID order by FS_News.id desc"
					Set NextRs = Conn.Execute(NextSql)
					If NextRs.eof or NextRs.bof Then
						  FS_PreviousTempStr = "没有了"
					Else
						if Application("UseDatePath")="1" then 
							FS_PreviousTempStr = "<a href='../.." & NextRs("path") & "/" &  NextRs("FileName") & "." & NextRs("FileExtName") & "' title ='"&NextRs("Title")&"'>"&NextRs("Title")&"</a>"
						else
							FS_PreviousTempStr = "<a href='" & NextRs("FileName") & "." & NextRs("FileExtName") & "' title ='"&NextRs("Title")&"'>"&NextRs("Title")&"</a>"	
						end if 
					End If 
					  NextRs.Close
					  Set NextRs = nothing
				End If
				If Instr(FileContent,"{%=NextPageNews()%}")<>0 then 
					'下一篇
					NextSql = "Select TOP 1 FS_news.title,FS_news.Path,FS_news.FileName,FS_news.FileExtName,FS_NewsClass.ClassEName From FS_News,FS_newsclass where FS_News.HeadNewsTF=0 and FS_News.DelTF=0 and FS_News.ID > " & NowNewsID & " and FS_News.ClassID = '" & NowClassID & "' and FS_News.ClassID=FS_NewsClass.ClassID order by FS_News.id"
					Set NextRs = Conn.Execute(NextSql)
					If NextRs.eof or NextRs.bof Then
						  FS_NextTempStr = "没有了"
					Else
						if Application("UseDatePath")="1" then 
							FS_NextTempStr = "<a href='../.." & NextRs("path") & "/" &  NextRs("FileName") & "." & NextRs("FileExtName") & "' title ='"&NextRs("Title")&"'>"&NextRs("Title")&"</a>"
						else
							FS_NextTempStr = "<a href='" & NextRs("FileName") & "." & NextRs("FileExtName") & "' title ='"&NextRs("Title")&"'>"&NextRs("Title")&"</a>"	
						end if
					End If
					  NextRs.Close
					  Set NextRs = nothing
				EnD If
				FileContent =replace(Replace(FileContent,"{%=PrePageNews()%}",FS_PreviousTempStr),"{%=NextPageNews()%}",FS_NextTempStr)
				'添加上一篇，下一篇功能 modi 2005-07-15
				'====================================
				FileContent = ReplaceAllServerFlag(FileContent)
			else
				FileContent = "模板内容为空"
			end if
		end if
		Set FileStreamObj = Nothing
		Set FileObj = Nothing
		Set FSOObj = Nothing
		'---------------------------/l
		CheckFolderExists TempSysRootDir & RsClassObj("SaveFilePath"),RsClassObj("ClassEName"),NewsRecordSetObj("Path") ,NewsRecordSetObj("FileName"),Application("UseDatePath")
		'---------------------------
		NewsContent = NewsRecordSetObj("Content")
		if IsNull(NewsContent) then NewsContent = ""
		if NewsContent <> "" and NewsRecordSetObj("LinkTF")="1" then NewsContent = ReplaceInnerLink(NewsContent)
		ContentArray = Split(NewsContent,"[Page]")
		NewsPageCount = UBound(ContentArray) + 1
		for i = LBound(ContentArray) to UBound(ContentArray)
			CurrPageNum = i + 1
			if NewsPageCount > 1 then
				NewsPageStr = "<p><div align=""right"">本新闻共<font color=red>" & NewsPageCount & "</font>页,当前在第<font color=red>" & CurrPageNum & "</font>页&nbsp;&nbsp;"
				for j = 1 to NewsPageCount
					if j = 1 then
						if CurrPageNum = j then
							NewsPageStr = NewsPageStr & "<font color=""red"">" & j & "</font>&nbsp;&nbsp;"
						else
							NewsPageStr = NewsPageStr & "<a href=""" & NewsRecordSetObj("FileName") & "." & NewsRecordSetObj("FileExtName") & """>" & j & "</a>&nbsp;&nbsp;"
						end if
					else
						if CurrPageNum = j then
							NewsPageStr = NewsPageStr & "<font color=""red"">" & j & "</font>&nbsp;&nbsp;"
						else
							NewsPageStr = NewsPageStr & "<a href="""  & NewsRecordSetObj("FileName") & "_" & j & "." & NewsRecordSetObj("FileExtName") & """>" & j & "</a>&nbsp;&nbsp;"
						end if
					end if				
				Next
				NewsPageStr = NewsPageStr & "</div></p>"
			else
				NewsPageStr = ""
			end if
			TempFileContent = FileContent
			SaveNewsContent = GetNewsContent(TempFileContent,NewsRecordSetObj,ContentArray(i) & NewsPageStr)
			if CurrPageNum = 1 then
				FileName = NewsRecordSetObj("FileName") & "." & NewsRecordSetObj("FileExtName")
			else
				FileName = NewsRecordSetObj("FileName") & "_" & CurrPageNum & "." & NewsRecordSetObj("FileExtName")
			end if
			if RsClassObj("SaveFilePath") = "/" then
			'-----------------------------------------------/l
			'laeep 判断是否使用日期路径
				if Application("UseDatePath")="0" then
					SaveFilePath = TempSysRootDir & RsClassObj("SaveFilePath") & RsClassObj("ClassEName") & "/" & FileName
				else
					SaveFilePath = TempSysRootDir & RsClassObj("SaveFilePath") & RsClassObj("ClassEName") & NewsRecordSetObj("Path") & "/" & FileName
				end if
			else
				if Application("UseDatePath")="0" then
					SaveFilePath = TempSysRootDir & RsClassObj("SaveFilePath") & "/" & RsClassObj("ClassEName") & "/" & FileName
				else
					SaveFilePath = TempSysRootDir & RsClassObj("SaveFilePath") & "/" & RsClassObj("ClassEName") & NewsRecordSetObj("Path") & "/" & FileName
				end if
			'-----------------------------------------------
			end if

			Select Case AvailableRefreshType
				Case 0
					FSOSaveFile SaveNewsContent,SaveFilePath
				Case 1
					SaveFile SaveNewsContent,SaveFilePath
				Case Else
					FSOSaveFile SaveNewsContent,SaveFilePath
			End Select
		Next
	end if
	Set RsClassObj = Nothing
End Function
Function ReplaceInnerLink(NewsContent)
	Dim RoutineSql,RsRoutineObj
	RoutineSql = "Select * from FS_Routine where Type=5"
	Set RsRoutineObj = Conn.Execute(RoutineSql)
	Dim StrReplace,Inti,DLocation,XLocation
	do while Not RsRoutineObj.Eof
		Inti=1
		StrReplace=RsRoutineObj("Name")
		If instr(1,NewsContent,StrReplace) then
			do while instr(Inti,NewsContent,StrReplace)<>0
				Inti=instr(Inti,NewsContent,StrReplace)
				'response.write Inti & "|"
				If Inti<>0 then
					DLocation=instr(Inti,NewsContent,">")'如果内容在><之间则替换
					XLocation=instr(Inti,NewsContent,"<")
					If DLocation>XLocation Then
						If instr(1,"[NoPage]",StrReplace)=0 then'避免替换[Page]里面的内容后，造成分页混乱
							NewsContent=left(NewsContent,Inti-1) & "<a href=" & RsRoutineObj("Url") & " target=_blank>" & StrReplace & "</a>" & mid(NewsContent,Inti+len(StrReplace))
							Inti=Inti+len("<a href=" & RsRoutineObj("Url") & " target=_blank>" & StrReplace & "</a>")
						Else
							Inti=Inti+len(StrReplace)
						End If
					Else
						Inti=Inti+len(StrReplace)
					end If
				End If
			
			loop
		End If
		RsRoutineObj.MoveNext
	Loop
	Set RsRoutineObj = Nothing
	ReplaceInnerLink = NewsContent
End Function
Function GetNewsContent(TempletContent,NewsRecordSet,NewsContent)
	TempletContent = Replace(TempletContent,"{News_Title}",NewsRecordSet("Title"))
	if Not IsNull(NewsRecordSet("SubTitle")) then
		TempletContent = Replace(TempletContent,"{News_SubTitle}",NewsRecordSet("SubTitle"))
	else
		TempletContent = Replace(TempletContent,"{News_SubTitle}","")
	end if
	if Not IsNull(NewsRecordSet("Author")) then
		TempletContent = Replace(TempletContent,"{News_Author}",NewsRecordSet("Author"))
	else
		TempletContent = Replace(TempletContent,"{News_Author}","")
	end if
	TempletContent = Replace(TempletContent,"{News_Content}",NewsContent)
	if Not IsNull(NewsRecordSet("TxtSource")) then
		TempletContent = Replace(TempletContent,"{News_TxtSource}",NewsRecordSet("TxtSource"))
	else
		TempletContent = Replace(TempletContent,"{News_TxtSource}","")
	end if
	if Not IsNull(NewsRecordSet("Editer")) then
		TempletContent = Replace(TempletContent,"{News_TxtEditer}",NewsRecordSet("Editer"))
	else
		TempletContent = Replace(TempletContent,"{News_TxtEditer}","")
	end if
	if Not IsNull(NewsRecordSet("AddDate")) then 
		TempletContent = Replace(TempletContent,"{News_AddDate}",NewsRecordSet("AddDate"))
	else
		TempletContent = Replace(TempletContent,"{News_AddDate}","")
	end if
	TempletContent = Replace(TempletContent,"{News_SendFriend}","<a href=" & AvailableDoMain & "/" & "Sendmail.asp?NewsID=" & NewsRecordSet("NewsID") & "  target=""_blank"">发送给好友</a>")
	TempletContent = Replace(TempletContent,"{News_ClickNum}","<script src=" & AvailableDoMain & "/" & "Click.asp?NewsID="& RefreshID &"></script>")
	TempletContent = Replace(TempletContent,"{News_ReviewContent}","<script src=" & AvailableDoMain & "/" & "ReviewContent.asp?NewsID="& NewsRecordSet("NewsID") &"></script>")
	'Added By Koolls at 2005.10.11
	TempletContent = Replace(TempletContent,"{News_Favorite}","<a target=""_blank"" Href=" & AvailableDoMain & "/" & UserDir &"/AddFavorite.asp?NewsID="& NewsRecordSet("ID") &">添加到收藏夹</a>")
	Dim ReviewStr
	if NewsRecordSet("ReviewTF") = 1 then
		ReviewStr = "<table width=""100%"" border=""0"" cellpadding=""3"" cellspacing=""1""><form name=""form1"" method=""post"" action=""" & AvailableDoMain & "/" & "NewsReview.asp?action=add&NewsID=" & NewsRecordSet("NewsID") & """><tr>"
		ReviewStr = ReviewStr & "<td width=""21%""><div align=right>会员名称：</div></td>"
		ReviewStr = ReviewStr & "<td width=""79%""> <input name=""MemName"" type=""text"" id=""MemName"" size=""10"" value="""">密码：<input name=""Password"" type=""password"" size=""8"" id=""Password""><input name=""NoName"" type=""checkbox"" id=""NoName"" value=""1"">匿名 <font color=""#FF0000"">·</font><a href=""" & AvailableDoMain & "/"& UserDir &"/sRegister.asp""><font color=""#FF0000"">注册</font></a>·<a href=""" & AvailableDoMain & "/"& UserDir &"/User_GetPassword.asp"">忘记密码？</a></td></tr>" 
		ReviewStr = ReviewStr & "<td>  <input name=""NewsID"" type=""hidden"" id=""NewsID"" value=""" & NewsRecordSet("NewsID") & """>"
		ReviewStr = ReviewStr & "<input name=""action"" type=""hidden"" id=""action"" value=""add""></tr>"
		ReviewStr = ReviewStr & "<tr><td> <div align=""right"">评论内容：<br>(最多300个字符) </div></td><td> <textarea name=""RevContent"" cols=""40"" rows=""5"" id=""RevContent""></textarea></td></tr>"
		ReviewStr = ReviewStr & "<tr><td></td><td> <input type=""submit"" name=""Submit"" value=""发表"">&nbsp;&nbsp;<a href=""" & AvailableDoMain & "/" & "NewsReview.asp?NewsID=" & NewsRecordSet("NewsID") & """><font color=red><b>查看评论</b></font></a></td></tr></form></table>"
	else
		ReviewStr = ""
	end if
	TempletContent = Replace(TempletContent,"{News_Review}",ReviewStr)
	GetNewsContent = TempletContent
End Function
Function RefreshDownLoad(DownLoadRecordSetObj) '生成下载
	Dim SaveFilePath,FileContent,TempletFileName,RsClassObj,SaveNewsContent,FSOObj,FileStreamObj,FileObj,FileName
	Dim TempSysRootDir
	SetRefreshValue "DownLoad",DownLoadRecordSetObj("DownLoadID")
	GetAvailableDoMain
	if SysRootDir = "" then
		TempSysRootDir = ""
	else
		TempSysRootDir = "/" & SysRootDir
	end if
	Set RsClassObj = Conn.Execute("Select ClassEName,SaveFilePath from FS_NewsClass where ClassID='" & DownLoadRecordSetObj("ClassID") & "'")
	if Not RsClassObj.Eof then
		TempletFileName = Server.MapPath(TempSysRootDir & DownLoadRecordSetObj("NewsTemplet"))
		Set FSOObj = Server.CreateObject(G_FS_FSO)
		if FSOObj.FileExists(TempletFileName) = False then
			FileContent = "模板不存在，请添加模板后再生成！"
		else
			Set FileObj = FSOObj.GetFile(TempletFileName)
			Set FileStreamObj = FileObj.OpenAsTextStream(1)
			if Not FileStreamObj.AtEndOfStream then
				FileContent = FileStreamObj.ReadAll
				if (DownLoadRecordSetObj("BrowPop") <> 0) And (DownLoadRecordSetObj("FileExtName") = "asp") then
					FileContent = GetPopStr(DownLoadRecordSetObj("BrowPop"),RsClassObj("SaveFilePath")) & FileContent
				end if
				FileContent = ReplaceAllServerFlag(ReplaceAllLable(FileContent))
			else
				FileContent = "模板内容为空"
			end if
		end if
		Set FileStreamObj = Nothing
		Set FileObj = Nothing
		Set FSOObj = Nothing
		'/l
		CheckFolderExists TempSysRootDir & RsClassObj("SaveFilePath"),RsClassObj("ClassEName"),"",DownLoadRecordSetObj("FileName"),"0"
		FileName = DownLoadRecordSetObj("FileName") & "." & DownLoadRecordSetObj("FileExtName")
		if RsClassObj("SaveFilePath") = "/" then
			SaveFilePath = TempSysRootDir & RsClassObj("SaveFilePath") & RsClassObj("ClassEName") & "/" & FileName
		else
			SaveFilePath = TempSysRootDir & RsClassObj("SaveFilePath") & "/" & RsClassObj("ClassEName") & "/" & FileName
		end if
		FileContent = GetDownLoadContent(FileContent,DownLoadRecordSetObj)
		Select Case AvailableRefreshType
			Case 0
				FSOSaveFile FileContent,SaveFilePath
			Case 1
				SaveFile FileContent,SaveFilePath
			Case Else
				FSOSaveFile FileContent,SaveFilePath
		End Select
	end if
	Set RsClassObj = Nothing
End Function
Function GetDownLoadContent(TempletContent,DownLoadRecordObj)
	Dim TempStr,AddressSql,RsAddressObj,DownLoadID,AddressStr,ReviewStr
	if Not DownLoadRecordObj.Eof then
		DownLoadID = DownLoadRecordObj("DownLoadID")
		if Not IsNull(DownLoadID) then
			AddressSql = "Select * from FS_DownLoadAddress where DownLoadID='" & DownLoadID & "'"
			Set RsAddressObj = Conn.Execute(AddressSql)
			if Not RsAddressObj.Eof then
				AddressStr = "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">" & Chr(13)
				do while Not RsAddressObj.Eof
					AddressStr = AddressStr & "<tr>" & Chr(13)
					AddressStr = AddressStr & "<td>" & Chr(13)
					AddressStr = AddressStr & "<a href=""" & AvailableDoMain & "/Down.asp?ID=" & RsAddressObj("ID") & "&DownID=" & RsAddressObj("DownLoadID") & """>点击下载--" & RsAddressObj("AddressName") & "</a>"
					AddressStr = AddressStr & "</td>" & Chr(13)
					AddressStr = AddressStr & "</tr>" & Chr(13)
					RsAddressObj.MoveNext
				Loop
				AddressStr = AddressStr & "</table>" & Chr(13)
			else
				AddressStr = ""
			end if
			Set RsAddressObj = Nothing
			TempletContent = Replace(TempletContent,"{DownLoad_Address}",AddressStr)
		else
			TempletContent = Replace(TempletContent,"{DownLoad_Address}","")
		end if
		if Not IsNull(DownLoadRecordObj("Name")) then
			TempletContent = Replace(TempletContent,"{DownLoad_Name}",DownLoadRecordObj("Name"))
		else
			TempletContent = Replace(TempletContent,"{DownLoad_Name}","")
		end if
		if Not IsNull(DownLoadRecordObj("Version")) then
			TempletContent = Replace(TempletContent,"{DownLoad_Version}",DownLoadRecordObj("Version"))
		else
			TempletContent = Replace(TempletContent,"{DownLoad_Version}","")
		end if
		TempletContent = Replace(TempletContent,"{DownLoad_ClickNum}","<script src=" & AvailableDoMain & "/" & "DownClick.asp?DownLoadID="& DownLoadRecordObj("DownLoadID") &"></script>")
		if Not IsNull(DownLoadRecordObj("Types")) then
			Select Case DownLoadRecordObj("Types")
				Case 1 TempStr = "图片"
				Case 2 TempStr = "文件"
				Case 3 TempStr = "程序"
				Case 4 TempStr = "Flash"
				Case 5 TempStr = "音乐"
				Case 6 TempStr = "影视"
				Case 7 TempStr = "其他"
				Case Else TempStr = ""
			End Select
			TempletContent = Replace(TempletContent,"{DownLoad_Types}",TempStr)
		else
			TempletContent = Replace(TempletContent,"{DownLoad_Types}","")
		end if
		if Not IsNull(DownLoadRecordObj("Language")) then
			Select Case DownLoadRecordObj("Language")
				Case 1 TempStr = "简体中文"
				Case 2 TempStr = "繁体中文"
				Case 3 TempStr = "英文"
				Case 4 TempStr = "法文"
				Case 5 TempStr = "日文"
				Case 6 TempStr = "德文"
				Case Else TempStr = ""
			End Select
			TempletContent = Replace(TempletContent,"{DownLoad_Language}",TempStr)
		else
			TempletContent = Replace(TempletContent,"{DownLoad_Language}","")
		end if
		if Not IsNull(DownLoadRecordObj("Accredit")) then
			Select Case DownLoadRecordObj("Accredit")
				Case 1 TempStr = "免费"
				Case 2 TempStr = "共享"
				Case 3 TempStr = "试用"
				Case 4 TempStr = "演示"
				Case 5 TempStr = "注册"
				Case 6 TempStr = "破解"
				Case 7 TempStr = "零售"
				Case 8 TempStr = "其他"
				Case Else TempStr = ""
			End Select
			TempletContent = Replace(TempletContent,"{DownLoad_Accredit}",TempStr)
		else
			TempletContent = Replace(TempletContent,"{DownLoad_Accredit}","")
		end if
		if Not IsNull(DownLoadRecordObj("FileSize")) then
			TempletContent = Replace(TempletContent,"{DownLoad_FileSize}",DownLoadRecordObj("FileSize"))
		else
			TempletContent = Replace(TempletContent,"{DownLoad_FileSize}","")
		end if
		if Not IsNull(DownLoadRecordObj("Appraise")) then
			Select Case DownLoadRecordObj("Appraise")
				Case 1 TempStr = "★"
				Case 2 TempStr = "★★"
				Case 3 TempStr = "★★★"
				Case 4 TempStr = "★★★★"
				Case 5 TempStr = "★★★★★"
				Case 6 TempStr = "★★★★★★"
				Case Else TempStr = ""
			End Select
			TempletContent = Replace(TempletContent,"{DownLoad_Appraise}",TempStr)
		else
			TempletContent = Replace(TempletContent,"{DownLoad_Appraise}","")
		end if
		if Not IsNull(DownLoadRecordObj("SystemType")) then
			TempletContent = Replace(TempletContent,"{DownLoad_SystemType}",DownLoadRecordObj("SystemType"))
		else
			TempletContent = Replace(TempletContent,"{DownLoad_SystemType}","")
		end if
		if DownloadRecordObj("ShowReviewTF") = 1 and  DownloadRecordObj("ReviewTF") = 1 Then
			TempletContent = Replace(TempletContent,"{DownLoad_ReviewContent}","<script src=" & AvailableDoMain & "/" & "ReviewContent.asp?DownloadID="& DownloadRecordObj("downloadid") &"></script>")
		else
			TempletContent = Replace(TempletContent,"{DownLoad_ReviewContent}","")
		end if
		if DownloadRecordObj("ReviewTF") = 1 then
			ReviewStr = "<table width=""100%"" border=""0"" cellpadding=""3"" cellspacing=""1""><form name=""form1"" method=""post"" action=""" & AvailableDoMain & "/" & "NewsReview.asp?action=add&DownloadID=" & DownloadRecordObj("downloadID") & """><tr>"
			ReviewStr = ReviewStr & "<td width=""21%""><div align=right>会员名称：</div></td>"
			ReviewStr = ReviewStr & "<td width=""79%""> <input name=""MemName"" type=""text"" id=""MemName"" size=""10"" value="""">密码：<input name=""Password"" type=""password"" size=""8"" id=""Password""><input name=""NoName"" type=""checkbox"" id=""NoName"" value=""1"">匿名<font color=""#FF0000"">·</font><a href=""" & AvailableDoMain & "/Users/Reg.asp""><font color=""#FF0000"">注册</font></a>·<a href=""" & AvailableDoMain & "/Users/UserForGet.asp"">忘记密码？</a></td></tr>" 
			ReviewStr = ReviewStr & "<td>  <input name=""DownloadID"" type=""hidden"" id=""DownloadID"" value=""" & DownloadRecordObj("downloadID") & """>"
			ReviewStr = ReviewStr & "<input name=""action"" type=""hidden"" id=""action"" value=""add""></tr>"
			ReviewStr = ReviewStr & "<tr><td> <div align=""right"">评论内容：<br>(最多300个字符) </div></td><td> <textarea name=""RevContent"" cols=""40"" rows=""5"" id=""RevContent""></textarea></td></tr>"
			ReviewStr = ReviewStr & "<tr><td></td><td> <input type=""submit"" name=""Submit"" value=""发表"">&nbsp;&nbsp;<a href=""" & AvailableDoMain & "/" & "NewsReview.asp?DownloadID=" & DownloadRecordObj("downloadID") & """><font color=red><b>查看评论</b></font></a></td></tr></form></table>"
		else
			ReviewStr = ""
		end if
		TempletContent = Replace(TempletContent,"{DownLoad_Review}",ReviewStr)
		if Not IsNull(DownLoadRecordObj("EMail")) then
			TempletContent = Replace(TempletContent,"{DownLoad_EMail}",DownLoadRecordObj("EMail"))
		else
			TempletContent = Replace(TempletContent,"{DownLoad_EMail}","")
		end if
		if Not IsNull(DownLoadRecordObj("ProviderUrl")) then
			TempletContent = Replace(TempletContent,"{DownLoad_ProviderUrl}",DownLoadRecordObj("ProviderUrl"))
		else
			TempletContent = Replace(TempletContent,"{DownLoad_ProviderUrl}","")
		end if
		if Not IsNull(DownLoadRecordObj("Provider")) then
			TempletContent = Replace(TempletContent,"{DownLoad_Provider}",DownLoadRecordObj("Provider"))
		else
			TempletContent = Replace(TempletContent,"{DownLoad_Provider}","")
		end if
		if Not IsNull(DownLoadRecordObj("PassWord")) then
			TempletContent = Replace(TempletContent,"{DownLoad_PassWord}",DownLoadRecordObj("PassWord"))
		else
			TempletContent = Replace(TempletContent,"{DownLoad_PassWord}","")
		end if
		if Not IsNull(DownLoadRecordObj("AddTime")) then
			TempletContent = Replace(TempletContent,"{DownLoad_AddTime}",DownLoadRecordObj("AddTime"))
		else
			TempletContent = Replace(TempletContent,"{DownLoad_AddTime}","")
		end if
		if Not IsNull(DownLoadRecordObj("EditTime")) then
			TempletContent = Replace(TempletContent,"{DownLoad_EditTime}",DownLoadRecordObj("EditTime"))
		else
			TempletContent = Replace(TempletContent,"{DownLoad_EditTime}","")
		end if
		TempletContent = Replace(TempletContent,"{DownLoad_Property}",DownLoadRecordObj("Property"))
		TempStr = DownLoadRecordObj("Description")
		if Not IsNull(TempStr) then
			TempletContent = Replace(TempletContent,"{DownLoad_Description}",TempStr)
		else
			TempletContent = Replace(TempletContent,"{DownLoad_Description}","")
		end if
		'=======================================
		'补足下载图片的显示地址，没有时不显示
		if instr(1,TempletContent,"{DownLoad_Pic}") then 
			if Not IsNull(DownLoadRecordObj("Pic")) then
				TempletContent = Replace(TempletContent,"{DownLoad_Pic}",AvailableDoMain & DownLoadRecordObj("Pic"))
			else
				dim PicEnd,PicBegin
				PicEnd=instr(1,TempletContent,"{DownLoad_Pic}")+14
				PicBegin=InstrRev(TempletContent,"<img",PicEnd)
				TempletContent = Replace(TempletContent,mid(TempletContent,PicBegin,PicEnd-PicBegin+2),"")
			end if
		end if
		'=======================================
	else
		TempletContent = ""
	end if
	GetDownLoadContent = TempletContent
End Function
Function GetPopStr(BrowPop,ClassSaveFilePath)'权限浏览
	Dim TempArray,TempStr,i
	if Not IsNull(ClassSaveFilePath) then
		TempArray = Split(ClassSaveFilePath,"/")
		for i = LBound(TempArray) to UBound(TempArray)
			TempStr = TempStr & "../"
		Next
	else
		GetPopStr = ""
		Exit Function
	end if
	GetPopStr = "<% GetGroupID " & BrowPop & " %" & ">" & Chr(13) & Chr(10)
	GetPopStr = GetPopStr & "<!--#include file=""" & TempStr & "Inc/JudgeRead.asp"" -->" & Chr(13) & Chr(10)
End Function

Sub CheckFolderExists(Path,ClassEName,DateDir,FileName,UseDatePath) '检查目录
	Dim FSOObj,TempPath',FolderObj,FileObj,ItemObj
	Dim DateDirChar
	Set FSOObj = Server.CreateObject(G_FS_FSO)
	TempPath = Server.MapPath(Path)
	if FSOObj.FolderExists(TempPath) = False then FSOObj.CreateFolder(TempPath)
	If ClassEName <> "" then
		TempPath=TempPath & "\" & ClassEName
		If FSOObj.FolderExists(TempPath) = False then
			FSOObj.CreateFolder(TempPath)
		Else
			'Set FolderObj = FSOObj.GetFolder(TempPath)
			'Set FileObj = FolderObj.Files
			If FSOObj.FileExists(FileName &".htm") then FSOObj.DeleteFile TempPath & "\" & FileName &".htm"
			If FSOObj.FileExists(FileName &".html") then FSOObj.DeleteFile TempPath & "\" & FileName &".html"
			If FSOObj.FileExists(FileName &".shtm") then FSOObj.DeleteFile TempPath & "\" & FileName &".shtm"
			If FSOObj.FileExists(FileName &".shtml") then FSOObj.DeleteFile TempPath & "\" & FileName &".shtml"
			If FSOObj.FileExists(FileName &".asp") then FSOObj.DeleteFile TempPath & "\" & FileName &".asp"
'			for Each ItemObj in FileObj
'				if InStr(LCase(ItemObj.name),FileName) then
'					FSOObj.DeleteFile TempPath & "\" & ItemObj.name
'				end if
'			Next
		End if
	End If

	DateDir=replace(DateDir,"/","\")
	if DateDir <> "" and Lcase(DateDir)<>"nouse" then
		'------------------------------/l
		'建立年目录
		DateDirChar=left(DateDir,instrrev(DateDir,"\")-1)
		TempPath = TempPath & DateDirChar
		if FSOObj.FolderExists(TempPath) = False and UseDatePath="1" then
			FSOObj.CreateFolder(TempPath)
		End if
		'--------------------------
		'建立月日目录或删除已存在的新闻文件
		TempPath = TempPath & right(DateDir,len(DateDir)-instrrev(DateDir,"\")+1)
		if FSOObj.FolderExists(TempPath) = False and UseDatePath="1" then
			FSOObj.CreateFolder(TempPath)
		elseif FSOObj.FolderExists(TempPath) = true then 
			If FSOObj.FileExists(FileName &".htm") then FSOObj.DeleteFile TempPath & "\" & FileName &".htm"
			If FSOObj.FileExists(FileName &".html") then FSOObj.DeleteFile TempPath & "\" & FileName &".html"
			If FSOObj.FileExists(FileName &".shtm") then FSOObj.DeleteFile TempPath & "\" & FileName &".shtm"
			If FSOObj.FileExists(FileName &".shtml") then FSOObj.DeleteFile TempPath & "\" & FileName &".shtml"
			If FSOObj.FileExists(FileName &".asp") then FSOObj.DeleteFile TempPath & "\" & FileName &".asp"
		end if
	end if
	Set FSOObj = Nothing
End Sub

Function GetOneNewsLinkURL(NewsID)
	Dim DoMain,TempParentID,RsParentObj,RsDoMainObj,ReturnValue
	Dim CheckRootClassIndex,CheckRootClassNumber,TempClassSaveFilePath,RootSaveFilePath,RootTF,NewsClassSaveFilePath
	RootTF = False
	Dim NewsSql,RsNewsObj
	CheckRootClassNumber = 30
	ReturnValue = ""
	NewsSql = "Select *,FS_NewsClass.FileExtName as ClassFileExtName,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.AuditTF=1 and FS_News.NewsID='" & NewsID & "'"
	Set RsNewsObj = Conn.Execute(NewsSql)
	if RsNewsObj.Eof then
		Set RsNewsObj = Nothing
		GetOneNewsLinkURL = ""
		Exit Function
	else
		if RsNewsObj("HeadNewsTF") = 1 then
			ReturnValue = RsNewsObj("HeadNewsPath")
		else
			if RsNewsObj("ParentID") <> "0" then
				Set RsParentObj = Conn.Execute("Select SaveFilePath,ParentID,Domain from FS_NewsClass where ClassID='" & RsNewsObj("ParentID") & "'")
				if Not RsParentObj.Eof then
					CheckRootClassIndex = 1
					TempParentID = RsParentObj("ParentID")
					do while Not (TempParentID = "0")
						CheckRootClassIndex = CheckRootClassIndex + 1
						RsParentObj.Close
						Set RsParentObj = Nothing
						Set RsParentObj = Conn.Execute("Select SaveFilePath,ParentID,Domain from FS_NewsClass where ClassID='" & TempParentID & "'")
						if RsParentObj.Eof then
							Set RsParentObj = Nothing
							Set RsNewsObj = Nothing
							GetOneNewsLinkURL = ""
							Exit Function
						end if
						TempParentID = RsParentObj("ParentID")
						if CheckRootClassIndex > CheckRootClassNumber then TempParentID = "0" '防止死循环
					Loop
					DoMain = RsParentObj("DoMain")
					RootSaveFilePath = RsParentObj("SaveFilePath")
					Set RsParentObj = Nothing
				else
					Set RsParentObj = Nothing
					Set RsNewsObj = Nothing
					GetOneNewsLinkURL = ""
					Exit Function
				end if
			else
				DoMain = RsNewsObj("DoMain")
				RootTF = True
				RootSaveFilePath =RsNewsObj("SaveFilePath")
			end if
			'/////////////////////////////////////////////l
			dim NewsDatePath
			if Application("UseDatePath")="1" then NewsDatePath=RsNewsObj("Path") else NewsDatePath=""
			if (Not IsNull(DoMain)) And (DoMain <> "") then
				If Instr(lCase(DoMain),"http://") = 0 Then
					DoMain = "http://"&DoMain
				End if
				if RootTF = True then
					ReturnValue = DoMain & "/" & RsNewsObj("ClassEName") & NewsDatePath & "/" & RsNewsObj("FileName") & "." & RsNewsObj("NewsFileExtName")
				else
					NewsClassSaveFilePath = RsNewsObj("SaveFilePath")
					NewsClassSaveFilePath = Replace(NewsClassSaveFilePath,RootSaveFilePath,"")
					ReturnValue = DoMain & NewsClassSaveFilePath & "/" & RsNewsObj("ClassEName") & NewsDatePath & "/" & RsNewsObj("FileName") & "." & RsNewsObj("NewsFileExtName")
				end if
			else
				if RsNewsObj("SaveFilePath") = "/" then
					TempClassSaveFilePath = RsNewsObj("SaveFilePath")
				else
					TempClassSaveFilePath = RsNewsObj("SaveFilePath") & "/"
				end if
				ReturnValue = AvailableDoMain & TempClassSaveFilePath & RsNewsObj("ClassEName") & NewsDatePath & "/" & RsNewsObj("FileName") & "." & RsNewsObj("NewsFileExtName")
			end if
			'/////////////////////////////////////////////
		end if
	end if
	Set RsNewsObj = Nothing
	GetOneNewsLinkURL = ReturnValue
End Function

Function GetOneDownLoadLinkURL(DownLoadID)
	Dim DoMain,TempParentID,RsParentObj,ReturnValue
	Dim DownLoadSql,RsDownLoadObj
	Dim CheckRootClassIndex,CheckRootClassNumber,TempClassSaveFilePath,RootTF,RootSaveFilePath,NewsClassSaveFilePath
	RootTF = False
	CheckRootClassNumber = 30
	ReturnValue = ""
	DownLoadSql = "Select *,FS_NewsClass.SaveFilePath,FS_NewsClass.FileExtName as ClassFileExtName,FS_Download.FileName,FS_DownLoad.FileExtName from FS_DownLoad,FS_NewsClass where FS_DownLoad.ClassID=FS_NewsClass.ClassID and FS_DownLoad.AuditTF=1 and FS_DownLoad.DownLoadID='" & DownLoadID & "'"
	Set RsDownLoadObj = Conn.Execute(DownLoadSql)
	if RsDownLoadObj.Eof then
		Set RsDownLoadObj = Nothing
		GetOneDownLoadLinkURL = ""
		Exit Function
	else
		if RsDownLoadObj("ParentID") <> "0" then
			Set RsParentObj = Conn.Execute("Select SaveFilePath,ParentID,Domain from FS_NewsClass where ClassID='" & RsDownLoadObj("ParentID") & "'")
			if Not RsParentObj.Eof then
				CheckRootClassIndex = 1
				TempParentID = RsParentObj("ParentID")
				do while Not (TempParentID = "0")
					CheckRootClassIndex = CheckRootClassIndex + 1
					RsParentObj.Close
					Set RsParentObj = Nothing
					Set RsParentObj = Conn.Execute("Select SaveFilePath,ParentID,Domain from FS_NewsClass where ClassID='" & TempParentID & "'")
					if RsParentObj.Eof then
						Set RsParentObj = Nothing
						Set RsDownLoadObj = Nothing
						GetOneDownLoadLinkURL = ""
						Exit Function
					end if
					TempParentID = RsParentObj("ParentID")
					if CheckRootClassIndex > CheckRootClassNumber then TempParentID = "0" '防止死循环
				Loop
				DoMain = RsParentObj("DoMain")
				RootSaveFilePath=RsParentObj("SaveFilePath")
				Set RsParentObj = Nothing
			else
				Set RsParentObj = Nothing
				Set RsDownLoadObj = Nothing
				GetOneDownLoadLinkURL = ""
				Exit Function
			end if
		else
			RootTF=True
			DoMain = RsDownLoadObj("DoMain")	
		end if
		if (Not IsNull(DoMain)) And (DoMain <> "") then
			If Instr(lCase(DoMain),"http://") = 0 Then
				DoMain = "http://"&DoMain
			End if
			if RootTF=true then 
				ReturnValue = DoMain & "/" & RsDownLoadObj("ClassEName") & "/" & RsDownLoadObj("FileName") & "." & RsDownLoadObj("FileExtName")
			else
					NewsClassSaveFilePath = RsDownLoadObj("SaveFilePath")
					NewsClassSaveFilePath = Replace(lcase(NewsClassSaveFilePath),lcase(RootSaveFilePath),"")
					ReturnValue = DoMain & NewsClassSaveFilePath & "/" & RsDownLoadObj("ClassEName") & "/" & RsDownLoadObj("FileName") & "." & RsDownLoadObj("FileExtName")
			end if
		else
			if RsDownLoadObj("SaveFilePath") = "/" then
				TempClassSaveFilePath = RsDownLoadObj("SaveFilePath")
			else
				TempClassSaveFilePath = RsDownLoadObj("SaveFilePath") & "/"
			end if
			ReturnValue = AvailableDoMain & TempClassSaveFilePath & RsDownLoadObj("ClassEName") & "/" & RsDownLoadObj("FileName") & "." & RsDownLoadObj("FileExtName")
		end if
	end if
	Set RsDownLoadObj = Nothing
	GetOneDownLoadLinkURL = ReturnValue
End Function

Function GetOneClassLinkURLByID(ClassID)
	Dim RsClassObj
	Set RsClassObj = Conn.Execute("Select SaveFilePath,ClassEName,FileExtName from FS_NewsClass where ClassID='" & ClassID & "'")
	GetOneClassLinkURLByID = GetOneClassLinkURL(RsClassObj("ClassEName"),RsClassObj("SaveFilePath"),RsClassObj("FileExtName"))
End Function

Function GetOneClassLinkURL(ClassEName,SaveFilePath,ClassFileExtName)
	Dim DoMain,TempParentID,RsParentObj,ReturnValue
	Dim CheckRootClassIndex,CheckRootClassNumber,TempClassSaveFilePath,RootTF,RootSaveFilePath
	RootTF = False
	CheckRootClassNumber = 30
	ReturnValue = ""
	Set RsParentObj = Conn.Execute("Select SaveFilePath,ParentID,Domain from FS_NewsClass where ClassEName='" & ClassEName & "'")
	if Not RsParentObj.Eof then
		if RsParentObj("ParentID") = "0" then
			DoMain = RsParentObj("DoMain")
			RootTF = True 
		else
			CheckRootClassIndex = 1
			TempParentID = RsParentObj("ParentID")
			do while Not (RsParentObj("ParentID") = "0")
				CheckRootClassIndex = CheckRootClassIndex + 1
				RsParentObj.Close
				Set RsParentObj = Nothing
				Set RsParentObj = Conn.Execute("Select SaveFilePath,ParentID,Domain from FS_NewsClass where ClassID='" & TempParentID & "'")
				if RsParentObj.Eof then
					Set RsParentObj = Nothing
					GetOneClassLinkURL = ""
					Exit Function
				end if
				TempParentID = RsParentObj("ParentID")
				if CheckRootClassIndex > CheckRootClassNumber then TempParentID = "0" '防止死循环 
			Loop
			DoMain = RsParentObj("DoMain")
			RootSaveFilePath = RsParentObj("SaveFilePath")
		end if 
	else
		Set RsParentObj = Nothing
		GetOneClassLinkURL = ""
		Exit Function
	end if
	Set RsParentObj = Nothing
	if (Not IsNull(DoMain)) And (DoMain <> "") then
		if RootTF = True then
			ReturnValue = "http://" & DoMain & "/" & ClassEName & "/index." & ClassFileExtName
		else
			SaveFilePath = Replace(SaveFilePath,RootSaveFilePath,"")
			ReturnValue = "http://" & DoMain & SaveFilePath & "/" & ClassEName & "/index." & ClassFileExtName
		end if
	else
		if SaveFilePath = "/" then
			TempClassSaveFilePath = SaveFilePath
		else
			TempClassSaveFilePath = SaveFilePath & "/"
		end if
		ReturnValue = AvailableDoMain & TempClassSaveFilePath & ClassEName & "/index." & ClassFileExtName
	end if
	GetOneClassLinkURL = ReturnValue
End Function

Function GetRowSpanNumber(DateRuleStr,DateRightStr,RowNumberStr)
	if DateRuleStr <> "" then
		if DateRightStr = "Left" then
			GetRowSpanNumber = "colspan=""" & RowNumberStr & """"
		elseif DateRightStr = "Center" then
			GetRowSpanNumber = "colspan=""" & RowNumberStr * 2 & """"
		elseif DateRightStr = "Right" then
			GetRowSpanNumber = "colspan=""" & RowNumberStr * 2 & """"
		else
			GetRowSpanNumber = "colspan=""" & RowNumberStr & """"
		end if
	else
		GetRowSpanNumber = "colspan=""" & RowNumberStr & """"
	end if
End Function

Function GetNewsNavitionStr(TxtNaviStr,NaviPicStr)
	if TxtNaviStr <> "" then
		GetNewsNavitionStr = TxtNaviStr
	else
		if NaviPicStr <> "" then 
			GetNewsNavitionStr = "<img src=""" & AvailableDoMain & NaviPicStr & """>"
		else
			GetNewsNavitionStr = ""
		end if
	end if
End Function

Function GetOpenTypeStr(OpenTypeStr)
	if OpenTypeStr = "1" then
		GetOpenTypeStr = " target=""_blank"""
	else
		GetOpenTypeStr = " "
	end if
End Function

Function GetTitleNumberStr(TitleNumber)
	If TitleNumber <> "" then
		GetTitleNumberStr = Cint(TitleNumber)
	Else
		GetTitleNumberStr = 10
	End If
End Function

Function GetCompatPicStr(CompatPicStr,DateRightStr,DateRuleStr,RowNumberStr)
	if CompatPicStr <> "" then
		if DateRightStr <> "" then
			CompatPicStr = "<tr>" & Chr(13) & Chr(10) & "<td Height=1 " & GetRowSpanNumber(DateRuleStr,DateRightStr,RowNumberStr) & ">" & Chr(13) & Chr(10) & "<table width=""100%"" cellpadding=""0"" cellspacing=""0"">" & Chr(13) & Chr(10) & "<tr>" & Chr(13) & Chr(10) & "<td Height=1 background=""" & AvailableDoMain & CompatPicStr & """>" & Chr(13) & Chr(10) & "</td>" & Chr(13) & Chr(10) & "</tr>" & Chr(13) & Chr(10) & "</table>" & Chr(13) & Chr(10) & "</td>" & Chr(13) & Chr(10) & "</tr>"
		else
			CompatPicStr = "<tr>" & Chr(13) & Chr(10) & "<td Height=1 " & GetRowSpanNumber(DateRuleStr,DateRightStr,RowNumberStr) & ">" & Chr(13) & Chr(10) & "<table width=""100%"" cellpadding=""0"" cellspacing=""0"">" & Chr(13) & Chr(10) & "<tr>" & Chr(13) & Chr(10) & "<td Height=1 background=""" & AvailableDoMain & CompatPicStr & """>" & Chr(13) & Chr(10) & "</td>" & Chr(13) & Chr(10) & "</tr>" & Chr(13) & Chr(10) & "</table>" & Chr(13) & Chr(10) & "</td>" & Chr(13) & Chr(10) & "</tr>"
		end if
	end if
	GetCompatPicStr = CompatPicStr
End Function

Function GetCSSStyleStr(CSSStyleStr)
	if CSSStyleStr <> "" then
		GetCSSStyleStr = " Class=""" & CSSStyleStr & """"
	else
		GetCSSStyleStr = ""
	end if
End Function

Function GetRecordSearchForm()
	GetRecordSearchForm = GetRecordSearchForm & "<table width=""100%;"" border=""0""><tr>"
	GetRecordSearchForm = GetRecordSearchForm & "<form target=""_blank"" method=""POST"" action=""" & AvailableDoMain & "/RecordSearch.asp" & """ name=""Record_Search_Form""><td>"
	GetRecordSearchForm = GetRecordSearchForm & "&nbsp;&nbsp;&nbsp;&nbsp;<select name=""SearchYear"" size=""1""><option value="""" selected> 选择年份 </option><option value=""2005"">2005</option><option value=""2004"">2004</option>"
	GetRecordSearchForm = GetRecordSearchForm & "<option value=""2003"">2003</option><option value=""2002"">2002</option><option value=""2001"">2001</option>"
	GetRecordSearchForm = GetRecordSearchForm & "<option value=""2000"">2000</option><option value=""1999"">1999</option></select>"
	GetRecordSearchForm = GetRecordSearchForm & "&nbsp;&nbsp;<select name=""SearchMonth"" size=""1""><option value="""" selected> 选择月份 </option><option value=""1"">1</option><option value=""2"">2</option>"
	GetRecordSearchForm = GetRecordSearchForm & "<option value=""3"">3</option><option value=""4"">4</option><option value=""5"">5</option>"
	GetRecordSearchForm = GetRecordSearchForm & "<option value=""6"">6</option><option value=""7"">7</option><option value=""8"">8</option>"
	GetRecordSearchForm = GetRecordSearchForm & "<option value=""9"">9</option><option value=""10"">10</option><option value=""11"">11</option>"
	GetRecordSearchForm = GetRecordSearchForm & "<option value=""12"">12</option></select>"
	GetRecordSearchForm = GetRecordSearchForm & "&nbsp;&nbsp;<select name=""SearchDate"" size=""1""><option value="""" selected> 选择日期 </option><option value=""1"">1</option><option value=""2"">2</option>"
	GetRecordSearchForm = GetRecordSearchForm & "<option value=""3"">3</option><option value=""4"">4</option><option value=""5"">5</option>"
	GetRecordSearchForm = GetRecordSearchForm & "<option value=""6"">6</option><option value=""7"">7</option><option value=""8"">8</option>"
	GetRecordSearchForm = GetRecordSearchForm & "<option value=""9"">9</option><option value=""10"">10</option><option value=""11"">11</option>"
	GetRecordSearchForm = GetRecordSearchForm & "<option value=""12"">12</option><option value=""13"">13</option><option value=""14"">14</option>"
	GetRecordSearchForm = GetRecordSearchForm & "<option value=""15"">15</option><option value=""16"">16</option><option value=""17"">17</option>"
	GetRecordSearchForm = GetRecordSearchForm & "<option value=""18"">18</option><option value=""19"">19</option><option value=""20"">20</option>"
	GetRecordSearchForm = GetRecordSearchForm & "<option value=""21"">21</option><option value=""22"">22</option><option value=""23"">23</option>"
	GetRecordSearchForm = GetRecordSearchForm & "<option value=""24"">24</option><option value=""25"">25</option><option value=""26"">26</option>"
	GetRecordSearchForm = GetRecordSearchForm & "<option value=""27"">27</option><option value=""28"">28</option><option value=""29"">29</option>"
	GetRecordSearchForm = GetRecordSearchForm & "<option value=""30"">30</option><option value=""31"">31</option></select>"
	GetRecordSearchForm = GetRecordSearchForm & "&nbsp;&nbsp;<input type=""submit"" value=""查看当日归档新闻"">"
	GetRecordSearchForm = GetRecordSearchForm & "</td>"
	GetRecordSearchForm = GetRecordSearchForm & "</form>"
	GetRecordSearchForm = GetRecordSearchForm & "</tr></table>"
End Function

Function MoveNewsFile(IDList,SourceClassID,TargetClassID)
'如果IDList不为空，则SourceClass为1时IDLis为新闻ID，为2时则IDLis为下载
'如果IDlist为空时，则SourceClass为要转移的类的 ID
Dim SqlStr,RsSource,RsTarget
Dim FSO,FolderObj,FilesObj,FileObj
Dim SourceDir,TarGetDir,sRootDir,DatePathStr
Set FSO = Server.CreateObject(G_FS_FSO)
If SysRootDir<>"" then 
	sRootDir="/" & SysRootDir
Else
	sRootDir=""
End If
If IdList<>"" then 
	IDList=replace(IDList,"***","','")
	If SourceClassID="1" then 
		SqlStr="Select FS_NewsClass.ClassEName,FS_NewsClass.SaveFilePath,FS_News.Path,FS_News.FileName,FS_News.FileExtName from FS_News,FS_NewsClass where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.NewsID in('" & IDList & "')"
	Else
		SqlStr="Select FS_NewsClass.ClassEName,FS_NewsClass.SaveFilePath,FS_Download.FileName,FS_Download.FileExtName from FS_Download,FS_NewsClass where FS_Download.ClassID=FS_NewsClass.ClassID and FS_Download.DownloadID in('" & IDList & "')"
	End If
Else
	SqlStr="Select FS_NewsClass.ClassEName,FS_NewsClass.SaveFilePath,FS_News.Path,FS_News.FileName,FS_News.FileExtName from FS_News,FS_NewsClass where FS_News.ClassID=FS_NewsClass.ClassID and FS_NewsClass.ClassID='" & SourceClassID & "'"
End If
Set RsSource=Conn.ExeCute(SqlStr)
Set RsTarget=Conn.ExeCute("Select ClassEName,SaveFilePath From FS_NewsClass where ClassID='" & TargetClassID & "'")
Do while Not RsSource.eof
	'得到日期路径
	If Application("UseDatePath")="1" and SourceClassID="1" then DatePathStr=RsSource("Path") Else DatePathStr=""
	'源文件路径
	SourceDir=sRootDir & RsSource("SaveFilePath") & "/" & RsSource("ClassEName") & DatePathStr & "/" & RsSource("FileName") & "." & RsSource("FileExtName")
	'目标文件路径
	TarGetDir=sRootDir & RsTarget("SaveFilePath") & "/" & RsTarget("ClassEName") & DatePathStr & "/" & RsSource("FileName") & "." & RsSource("FileExtName")

	SourceDir=Server.MapPath(SourceDir)
	TarGetDir=Server.MapPath(TarGetDir)
	if (FSO.FileExists(SourceDir)) then
	'如果目录不存在，则创建目录
		CreatMoreDir TarGetDir,instr(TarGetDir,replace(RsTarget("SaveFilePath"),"/","\"))
		FSO.MoveFile SourceDir,TarGetDir
	End If
	RsSource.MoveNext
Loop
'--------------------------------
'合并栏目时，用来转移栏目中的下载
If IDList="" then 
	SqlStr="Select FS_NewsClass.ClassEName,FS_NewsClass.SaveFilePath,FS_Download.FileName,FS_Download.FileExtName from FS_Download,FS_NewsClass where FS_Download.ClassID=FS_NewsClass.ClassID and FS_NewsClass.ClassID='" & SourceClassID & "'"
	Set RsSource=Conn.ExeCute(SqlStr)
	Do while Not RsSource.eof
		'源文件路径
		SourceDir=sRootDir & RsSource("SaveFilePath") & "/" & RsSource("ClassEName") & "/" & RsSource("FileName") & "." & RsSource("FileExtName")
		'目标文件路径
		TarGetDir=sRootDir & RsTarget("SaveFilePath") & "/" & RsTarget("ClassEName") &  "/" & RsSource("FileName") & "." & RsSource("FileExtName")

		SourceDir=Server.MapPath(SourceDir)
		TarGetDir=Server.MapPath(TarGetDir)
		if (FSO.FileExists(SourceDir)) then
		'如果目录不存在，则创建目录
			CreatMoreDir TarGetDir,instr(TarGetDir,replace(RsTarget("SaveFilePath"),"/","\"))
			FSO.MoveFile SourceDir,TarGetDir
		End If
		RsSource.MoveNext
	Loop

End If
'------------------------------------
Set FSO = Nothing
Set RsSource = Nothing
Set RsTarget = Nothing
End Function

Function CreatMoreDir(DirStr,iBegin)
	Dim sBuild,sDir,FSO
	Set FSO = Server.CreateObject(G_FS_FSO)
	sBuild = left(DirStr,iBegin - 1)
	sDir = Mid(DirStr,iBegin)
	While InStr(2, sDir,"\") > 1
		sBuild = sBuild & left(sDir,InStr(2,sDir,"\") - 1)
		sDir = Mid(sDir,InStr(2,sDir,"\"))
		If (FSO.FolderExists(sBuild)) then
		else
			FSO.CreateFolder(sBuild)
		End IF
	Wend
	set FSO=Nothing
End Function 

Function AutoSplitPages(StrNewsContent)
Dim Inti,StrTrueContent,iPageLen,DLocation,XLocation,FoundStr
	If StrNewsContent<>"" and AutoPagesNum<>0 and instr(1,StrNewsContent,"[Page]")=0 then
		Inti=instr(1,StrNewsContent,"<")
		If inti>=1 then '新闻中存在Html标记
			StrTrueContent=left(StrNewsContent,Inti-1)
			iPageLen=IStrLen(StrTrueContent)
			inti=inti+1
		Else			'新闻中不存在Html标记，对内容直接分页即可
			dim i,c,t
			do while i< len(StrNewsContent)
			i=i+1
				c=Abs(Asc(Mid(StrNewsContent,i,1)))
				if c>255 then	'判断为汉字则为两个字符，英文为一个字符
					t=t+2
				else
					t=t+1
				end if
				if t>=AutoPagesNum then		'如果字数达到了分页的数量则插入分页符号
					StrNewsContent=left(StrNewsContent,i)&"[Page]"&mid(StrNewsContent,i+1)
					i=i+6
					t=0
				end if
			loop
			AutoSplitPages=StrNewsContent	'返回插入分页符号的内容
			Exit Function
		End If
		iPageLen=0
		'新闻中存在Html标记时，则用下面的语句来处理
		do while instr(Inti,StrNewsContent,">")<>0
			DLocation=instr(Inti,StrNewsContent,">")		'只计算Html标记之外的字符数量
			XLocation=instr(DLocation,StrNewsContent,"<")
			If XLocation>DLocation+1 then
				Inti=XLocation
				StrTrueContent=mid(StrNewsContent,DLocation+1,XLocation-DLocation-1)
				iPageLen=iPageLen+IStrLen(StrTrueContent)	'统计Html之外的字符的数量
				If iPageLen>AutoPagesNum then				'如果达到了分页的数量则插入分页字符
					FoundStr=Lcase(left(StrNewsContent,XLocation-1))
					If AllowSplitPages(FoundStr,"table|a|b>|i>|strong|div")=true then
						StrNewsContent=left(StrNewsContent,XLocation-1)&"[Page]"&mid(StrNewsContent,XLocation)
						iPageLen=0								'重新统计Html之外的字符
					End If
				End If
			ElseIf XLocation=0 then							'在后面再也找不到<，即后面没有Html标记了
				Exit Do
			ElseIf XLocation=DLocation+1 then				'找到的Html标记之间的内容为空，则继续向后找
				Inti=XLocation
			End If
		loop
	End If
AutoSplitPages=StrNewsContent
End Function

Function AllowSplitPages(TempStr,FindStr)
	Dim Inti,BeginStr,EndStr,BeginStrNum,EndStrNum,ArrStrFind,i
	If TempStr<>"" and FindStr<>"" then
		ArrStrFind=split(FindStr,"|")
		For i = 0 to Ubound(ArrStrFind)
			BeginStr="<"&ArrStrFind(i)
			EndStr  ="</"&ArrStrFind(i)
			Inti=0
			do while instr(Inti+1,TempStr,BeginStr)<>0
				Inti=instr(Inti+1,TempStr,BeginStr)
				BeginStrNum=BeginStrNum+1
			Loop
			Inti=0
			do while instr(Inti+1,TempStr,EndStr)<>0
				Inti=instr(Inti+1,TempStr,EndStr)
				EndStrNum=EndStrNum+1
			Loop
			If EndStrNum=BeginStrNum then
				AllowSplitPages=true
			Else
				AllowSplitPages=False
				Exit Function
			End If
		Next
	Else
		AllowSplitPages=False
	End If
End Function

Function WebDomain
Dim LocalPort
If Request.ServerVariables("SERVER_PORT")<>"80" Then
	LocalPort=":"&Request.ServerVariables("SERVER_PORT")
Else
	LocalPort=""
End If
WebDomain="http://"&Request.ServerVariables("SERVER_NAME")&LocalPort
End Function


'******************************
'把{FS_当前类目}转成类目名称
'author:lino
'Start
'*****************************
Function ReplaceYpren(Content)
Dim whatIsClass
whatIsClass=GetLableContent("ypren,")

Content=replace(Content,"{FS_当前类目}",whatIsCLass)
ReplaceYpren=Content
  'response.write("Content is:"&Content)
End Function

'**************************
'End
'**************************

%>