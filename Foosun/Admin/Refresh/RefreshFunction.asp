<%
Dim RefreshType,RefreshID,AvailableDoMain,AvailableRefreshType,Dummy_Riker
Dim NotReplaceLableArray  '需要分页的标签
Dim NotReplaceLableOldArray '需要分页的标签
NotReplaceLableArray = "" 
NotReplaceLableOldArray = ""  
If SysRootDir <> "" Then
	Dummy_Riker = "/" & SysRootDir
Else
	Dummy_Riker = ""
End If
Sub SetRefreshValue(RefreshTypeStr,RefreshIDStr)
	RefreshType = RefreshTypeStr
	RefreshID = RefreshIDStr
End Sub
Sub GetAvailableDoMain()
	Dim ConfigSql,RsConfigObj,ShopIsOpen
	ConfigSql = "Select IsShop,DoMain,MakeType,IndexExtName from FS_Config"
	Set RsConfigObj = Conn.Execute(ConfigSql)
	if Not RsConfigObj.Eof then
		AvailableDoMain = RsConfigObj("DoMain")
		if Not IsNull(RsConfigObj("MakeType")) then
			AvailableRefreshType = RsConfigObj("MakeType")
		else
			AvailableRefreshType = 0
		end if
		ShopIsOpen=RsConfigObj("IsShop")
	else
		AvailableDoMain = GetDoMain
		AvailableRefreshType = 0
		ShopIsOpen=0
	end if
	Set RsConfigObj = Nothing
End Sub
'归档新闻列表
Function LableFile(TitleNumberStr,CompatPicStr,NaviPicStr,DateRuleStr,DateRightStr,RowHeightStr,RowNumberStr,ShowClassCNNameStr,CSSStyleStr,OpenTypeStr,DateCSSStyleStr,TxtNaviStr) 
	if RefreshType <> "Record" then
		LableFile = ""
		Exit Function
	end if
	Dim i,RsClassObj,ClassName,TempDateShowStr
	OpenTypeStr = GetOpenTypeStr(OpenTypeStr)
	NaviPicStr = GetNewsNavitionStr(TxtNaviStr,NaviPicStr)
	CompatPicStr = GetCompatPicStr(CompatPicStr,DateRightStr,DateRuleStr,RowNumberStr)
	if RowHeightStr <> "" then RowHeightStr = " Height=""" & RowHeightStr & """"
	LableFile = "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">" & Chr(13) & Chr(10)
	LableFile =  LableFile & "<tr><td align=""center"" colspan=""" & RowNumberStr & """><font size=""5""><strong>" & RefreshTime & "归档新闻</strong></font></td></tr>" & Chr(13) & Chr(10)
	do while Not RsRecordObj.Eof
		LableFile = LableFile & "<tr " & RowHeightStr & ">" & Chr(13) & Chr(10)
		for i = 1 to RowNumberStr
			if DateRuleStr <> "" then
				if DateRightStr = "Left" then
					TempDateShowStr = "&nbsp;&nbsp;<span " & GetCSSStyleStr(DateCSSStyleStr) & ">" & DateFormat(RsRecordObj("AddDate"),DateRuleStr) & "</span>"
				elseif DateRightStr = "Center" then
					TempDateShowStr = "<td align=""center""><span " & GetCSSStyleStr(DateCSSStyleStr) & ">" & DateFormat(RsRecordObj("AddDate"),DateRuleStr) & "</span>" & "</td>"& Chr(13) & Chr(10)
				elseif DateRightStr = "Right" then
					TempDateShowStr = "<td align=""Right""><span " & GetCSSStyleStr(DateCSSStyleStr) & ">" & DateFormat(RsRecordObj("AddDate"),DateRuleStr) & "</span>" & "</td>" & Chr(13) & Chr(10)
				else
					TempDateShowStr = "&nbsp;&nbsp;<span " & GetCSSStyleStr(DateCSSStyleStr) & ">" & DateFormat(RsRecordObj("AddDate"),DateRuleStr) & "</span>"
				end if
			else
				TempDateShowStr = ""
			end if
			if ShowClassCNNameStr = "1" then
				ClassName = ""
				Set RsClassObj = Conn.Execute("Select * from FS_NewsClass where ClassID='" & RsRecordObj("ClassID") & "'")
				if Not RsClassObj.Eof then ClassName = "[" & RsClassObj("ClassCName") & "]"
				Set RsClassObj = Nothing
			end if
			if DateRightStr = "Center" Or DateRightStr = "Right" then
				LableFile = LableFile & "<td>" & NaviPicStr & ClassName & "<a " & OpenTypeStr & GetCSSStyleStr(CSSStyleStr) & " href=""" & GetRecordOneNewsLink(RsRecordObj) & """  title="""& RsRecordObj("Title")&""">" & GetHTMLTitle(RsRecordObj("TitleStyle"),GotTopic(RsRecordObj("Title"),TitleNumberStr)) & "</a></td>" & TempDateShowStr & Chr(13) & Chr(10)
			else
				LableFile = LableFile & "<td>" & NaviPicStr & ClassName & "<a " & OpenTypeStr & GetCSSStyleStr(CSSStyleStr) & " href=""" & GetRecordOneNewsLink(RsRecordObj) & """  title="""& RsRecordObj("Title")&""">" & GetHTMLTitle(RsRecordObj("TitleStyle"),GotTopic(RsRecordObj("Title"),TitleNumberStr)) & "</a>" & TempDateShowStr & "</td>" & Chr(13) & Chr(10)
			end if
			RsRecordObj.MoveNext
			if RsRecordObj.Eof then Exit For
		Next
		LableFile = LableFile & "</tr>" & Chr(13) & Chr(10)
		LableFile = LableFile & CompatPicStr
	Loop
	LableFile =  LableFile & "<tr><td height=""50"" align=""center"" colspan=""" & RowNumberStr & """>" & GetRecordSearchForm & "</td></tr>" & Chr(13) & Chr(10)
	LableFile = LableFile & "</table>" & Chr(13) & Chr(10)
End Function

Function GetRecordOneNewsLink(Obj)
	Dim DoMain,TempParentID,RsParentObj,ReturnValue,RsClassObj,LoopTF
	Dim CheckRootClassIndex,CheckRootClassNumber,TempClassSaveFilePath
	CheckRootClassNumber = 30
	LoopTF = False
	ReturnValue = ""
	if Obj("HeadNewsTF") = 1 then
		ReturnValue = Obj("HeadNewsPath")
	else
		Set RsClassObj = Conn.Execute("Select * from FS_NewsClass where ClassID='" & Obj("ClassID") & "'")
		if Not RsClassObj.Eof then
			Set RsParentObj = Conn.Execute("Select ParentID,Domain from FS_NewsClass where ClassID='" & Obj("ClassID") & "'")
			TempParentID = RsParentObj("ParentID")
			do while Not (TempParentID = "0")
				LoopTF = True
				CheckRootClassIndex = CheckRootClassIndex + 1
				RsParentObj.Close
				Set RsParentObj = Nothing
				Set RsParentObj = Conn.Execute("Select ParentID,Domain from FS_NewsClass where ClassID='" & TempParentID & "'")
				if RsParentObj.Eof then
					Set RsParentObj = Nothing
					Set RsClassObj = Nothing
					GetRecordOneNewsLink = ""
					Exit Function
				end if
				TempParentID = RsParentObj("ParentID")
				if CheckRootClassIndex > CheckRootClassNumber then TempParentID = "0" '防止死循环
			Loop
			if LoopTF = True then
				DoMain = RsParentObj("DoMain")
			else
				DoMain = RsClassObj("DoMain")
			end if
			Set RsParentObj = Nothing
			'=======================
			'归档文件是否使用日期路径判断
			dim NewsDatePath
			if Application("UseDatePath")="1" then NewsDatePath=Obj("Path") else NewsDatePath=""
			if (Not IsNull(DoMain)) And (DoMain <> "") then
				ReturnValue = "http://" & DoMain & "/" & RsClassObj("ClassEName")& NewsDatePath & "/" & Obj("FileName") & "." & Obj("FileExtName")
			else
				if RsClassObj("SaveFilePath") = "/" then
					TempClassSaveFilePath = RsClassObj("SaveFilePath")
				else
					TempClassSaveFilePath = RsClassObj("SaveFilePath") & "/"
				end if
				ReturnValue = AvailableDoMain & TempClassSaveFilePath & RsClassObj("ClassEName") &NewsDatePath& "/" & Obj("FileName") & "." & Obj("FileExtName")
			end if
			'=======================
		else
			ReturnValue = ""
		end if
		Set RsClassObj = Nothing
	end if
	GetRecordOneNewsLink = ReturnValue
End Function
'调用大栏目
Function SelfClass(ClassEName,NewsListNumberStr,TitleNumberStr,CompatPicStr,NaviPicStr,DateRuleStr,DateRightStr,RowHeightStr,RowNumberStr,ShowClassCNNameStr,MoreLinkTypeStr,MoreLinkContentStr,CSSStyleStr,OpenTypeStr,DateCSSStyleStr,TxtNaviStr)
	Dim RsNewsObj,NewsSql,RsClassObj,ClassSql,AllClassID,i,ClassCNName
	Dim TempDateShowStr,ReViewStr
	TitleNumberStr = GetTitleNumberStr(TitleNumberStr)
	OpenTypeStr = GetOpenTypeStr(OpenTypeStr)
	NaviPicStr = GetNewsNavitionStr(TxtNaviStr,NaviPicStr)
	if RowHeightStr <> "" then RowHeightStr = " Height=""" & RowHeightStr & """"
	CompatPicStr = GetCompatPicStr(CompatPicStr,DateRightStr,DateRuleStr,RowNumberStr)
	ClassSql = "Select ClassCName,ClassEName,ClassID,SaveFilePath,FileExtName from FS_NewsClass where ClassEName='" & ClassEName & "'"
	Set RsClassObj = Conn.Execute(ClassSql)
	if Not RsClassObj.Eof then
		AllClassID = "'" & RsClassObj("ClassID") & "'" & ChildClassIDList(RsClassObj("ClassID"))
		NewsSql = "Select top " & NewsListNumberStr & " *,FS_NewsClass.FileExtName as ClassFileExtName,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.AuditTF=1 and FS_News.delTF=0 and FS_NewsClass.ClassID in (" & AllClassID & ") order by FS_News.ID Desc"
		SelfClass = "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">" & Chr(13) & Chr(10)
		Set RsNewsObj = Conn.Execute(NewsSql)
		do while Not RsNewsObj.Eof
'新闻标题后面加评论
			If RsNewsObj("TitleShowReview")="1" then
				ReViewStr="    <a href="""&AvailableDoMain&"/NewsReview.asp?NewsID="&RsNewsObj("NewsID")&""">评论</a>"
			Else
				ReViewStr=""
			End If
			SelfClass = SelfClass & "<tr>" & Chr(13) & Chr(10)
			for i = 1 to RowNumberStr
				if DateRuleStr <> "" then
					if DateRightStr = "Left" then
						TempDateShowStr = "&nbsp;&nbsp;<span " & GetCSSStyleStr(DateCSSStyleStr) & ">" & DateFormat(RsNewsObj("AddDate"),DateRuleStr) & "</span>"
					elseif DateRightStr = "Center" then
						TempDateShowStr = "<td align=""center""><span " & GetCSSStyleStr(DateCSSStyleStr) & ">" & DateFormat(RsNewsObj("AddDate"),DateRuleStr) & "</span>" & "</td>"& Chr(13) & Chr(10)
					elseif DateRightStr = "Right" then
						TempDateShowStr = "<td align=""Right""><span " & GetCSSStyleStr(DateCSSStyleStr) & ">" & DateFormat(RsNewsObj("AddDate"),DateRuleStr) & "</span>" & "</td>" & Chr(13) & Chr(10)
					else
						TempDateShowStr = "&nbsp;&nbsp;<span " & GetCSSStyleStr(DateCSSStyleStr) & ">" & DateFormat(RsNewsObj("AddDate"),DateRuleStr) & "</span>"
					end if
				else
					TempDateShowStr = ""
				end if
				if ShowClassCNNameStr = "1" then
					ClassCNName = "<a " & OpenTypeStr & GetCSSStyleStr(CSSStyleStr) & " href=""" & GetOneClassLinkURL(RsNewsObj("ClassEName"),RsNewsObj("SaveFilePath"),RsNewsObj("ClassFileExtName")) & """ >[" & GotTopic(RsNewsObj("ClassCName"),TitleNumberStr) & "]</a>&nbsp;"
				else
					ClassCNName = ""
				end if
				if DateRightStr = "Center" Or DateRightStr = "Right" then
					SelfClass = SelfClass & "<td " & RowHeightStr & ">" & NaviPicStr & ClassCNName & "<a " & OpenTypeStr & GetCSSStyleStr(CSSStyleStr) & " href=""" & GetOneNewsLinkURL(RsNewsObj("NewsID")) & """  title="""& RsNewsObj("Title")&""">" & GetHTMLTitle(RsNewsObj("TitleStyle"),GotTopic(RsNewsObj("Title"),TitleNumberStr)) & "</a>" & ReViewStr & "</td>" & TempDateShowStr & Chr(13) & Chr(10)
				else
					SelfClass = SelfClass & "<td " & RowHeightStr & ">" & NaviPicStr & ClassCNName & "<a " & OpenTypeStr & GetCSSStyleStr(CSSStyleStr) & " href=""" & GetOneNewsLinkURL(RsNewsObj("NewsID")) & """  title="""& RsNewsObj("Title")&""">" & GetHTMLTitle(RsNewsObj("TitleStyle"),GotTopic(RsNewsObj("Title"),TitleNumberStr)) & "</a>" & ReViewStr & TempDateShowStr & "</td>" & Chr(13) & Chr(10)
				end if
				RsNewsObj.MoveNext
				if RsNewsObj.Eof then Exit For
			Next
			SelfClass = SelfClass & "</tr>" & Chr(13) & Chr(10)
			SelfClass = SelfClass & CompatPicStr
		Loop
		if MoreLinkContentStr <> "" then
			if MoreLinkTypeStr = "1" then
				MoreLinkContentStr="<a " & OpenTypeStr & GetCSSStyleStr(CSSStyleStr) & " href=""" & GetOneClassLinkURL(RsClassObj("ClassEName"),RsClassObj("SaveFilePath"),RsClassObj("FileExtName")) & """ ><img border=0 src=""" & AvailableDoMain & MoreLinkContentStr & """></a>"
			elseif MoreLinkTypeStr = "0" then
				MoreLinkContentStr = "<a " & OpenTypeStr & GetCSSStyleStr(CSSStyleStr) & " href=""" & GetOneClassLinkURL(RsClassObj("ClassEName"),RsClassObj("SaveFilePath"),RsClassObj("FileExtName")) & """ >" & MoreLinkContentStr & "</a>"
			else
				MoreLinkContentStr = ""
			end if
			if DateRuleStr <> "" then
				SelfClass = SelfClass & "<tr><td " & GetRowSpanNumber(DateRuleStr,DateRightStr,RowNumberStr) & " align=""right"">" & MoreLinkContentStr & "</td></tr>" & Chr(13) & Chr(10)
			else
				SelfClass = SelfClass & "<tr><td align=""right"" " & GetRowSpanNumber(DateRuleStr,DateRightStr,RowNumberStr) & ">" & MoreLinkContentStr & "</td></tr>" & Chr(13) & Chr(10)
			end if
		end if
		SelfClass = SelfClass & "</table>" & Chr(13) & Chr(10)
		Set RsNewsObj = Nothing
	else
		SelfClass = ""
	end if
	Set RsClassObj = Nothing
End Function
'调用栏目子栏目
Function ChildClassList(ClassNumberStr,NewsNumberStr,CompatPicStr,NaviPicStr,ClassRowHeightStr,NewsRowHeightStr,ClassRowNumberStr,NewsRowNumberStr,DateRuleStr,DateRightStr,TitleNumberStr,MoreLinkTypeStr,MoreLinkContentStr,ClassBGPicStr,CSSStyleStr,OpenTypeStr,DateCSSStyleStr,TxtNaviStr)
	Dim TempSetNewsRowHeightStr
	Dim TempSetNewsNumberStr
	Dim TempSetTitleNumberStr
	Dim TempSetCompatPicStr
	Dim TempSetNaviPicStr
	Dim TempSetDateRuleStr
	Dim TempSetDateRightStr
	Dim TempSetNewsRowNumberStr
	Dim TempSetMoreLinkTypeStr
	Dim TempSetMoreLinkContentStr
	Dim TempSetCSSStyleStr
	Dim TempSetOpenTypeStr
	Dim TempSetDateCSSStyleStr
	Dim TempSetTxtNaviStr
	TempSetNewsRowHeightStr = NewsRowHeightStr
	If TitleNumberStr <> "" then
		TitleNumberStr = Cint(TitleNumberStr)
	Else
		TitleNumberStr = 10
	End If
	if RefreshType = "Class" then
		Dim ClassSql,RsClassObj,AllChildClassID,i
		AllChildClassID = ChildClassIDList(RefreshID)
		if AllChildClassID <> "" then
			if Left(AllChildClassID,1) = "," then AllChildClassID = Right(AllChildClassID,Len(AllChildClassID)-1)
		else
			ChildClassList = ""
			Exit Function
		end if
		if ClassBGPicStr <> "" then
			ClassBGPicStr = "<tr>" & Chr(13) & Chr(10) & "<td Height=1 colspan=""" & ClassRowNumberStr & """>" & Chr(13) & Chr(10) & "<table width=""100%"" cellpadding=""0"" cellspacing=""0"">" & Chr(13) & Chr(10) & "<tr>" & Chr(13) & Chr(10) & "<td Height=1 background=""" & ClassBGPicStr & """>" & Chr(13) & Chr(10) & "</td>" & Chr(13) & Chr(10) & "</tr>" & Chr(13) & Chr(10) & "</table>" & Chr(13) & Chr(10) & "</td>" & Chr(13) & Chr(10) & "</tr>"
		end if
		ClassSql = "Select Top " & ClassNumberStr & " * from FS_NewsClass where ClassID in (" & AllChildClassID & ") and DelFlag=0 order by ID desc"
		Set RsClassObj = Conn.Execute(ClassSql)
		if Not RsClassObj.Eof then
			ChildClassList = "<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""100%"">" & Chr(13) & Chr(10)
			do while Not RsClassObj.Eof
				TempSetNewsRowHeightStr = NewsRowHeightStr
				TempSetNewsRowHeightStr = NewsRowHeightStr
				TempSetNewsNumberStr = NewsNumberStr
				TempSetTitleNumberStr = TitleNumberStr
				TempSetCompatPicStr = CompatPicStr
				TempSetNaviPicStr = NaviPicStr
				TempSetDateRuleStr = DateRuleStr
				TempSetDateRightStr = DateRightStr
				TempSetNewsRowNumberStr = NewsRowNumberStr
				TempSetMoreLinkTypeStr = MoreLinkTypeStr
				TempSetMoreLinkContentStr = MoreLinkContentStr
				TempSetCSSStyleStr = CSSStyleStr
				TempSetOpenTypeStr = OpenTypeStr
				TempSetDateCSSStyleStr = DateCSSStyleStr
				TempSetTxtNaviStr = TxtNaviStr
				ChildClassList = ChildClassList & "<tr>" & Chr(13) & Chr(10)
				For i = 1 to ClassRowNumberStr
					ChildClassList = ChildClassList & "<td valign=""top"">" & GetOneClassNewsList(RsClassObj,TempSetNewsNumberStr,TempSetTitleNumberStr,TempSetCompatPicStr,TempSetNaviPicStr,TempSetDateRuleStr,TempSetDateRightStr,TempSetNewsRowHeightStr,TempSetNewsRowNumberStr,TempSetMoreLinkTypeStr,TempSetMoreLinkContentStr,TempSetCSSStyleStr,TempSetOpenTypeStr,TempSetDateCSSStyleStr,TempSetTxtNaviStr) & "</td>"
					RsClassObj.MoveNext
					if RsClassObj.Eof then
						Exit For
					end if
				Next
				ChildClassList = ChildClassList & "</tr>" & Chr(13) & Chr(10)
				ChildClassList = ChildClassList & ClassBGPicStr & Chr(13) & Chr(10)
			Loop
			ChildClassList = ChildClassList & "</table>" & Chr(13) & Chr(10)
		else
			ChildClassList = ""
		end if
	else
		ChildClassList = ""
	end if
End Function
'得到一个栏目的新闻列表
Function GetOneClassNewsList(AlreadyClassObj,NewsListNumberStr,TitleNumberStr,CompatPicStr,NaviPicStr,DateRuleStr,DateRightStr,RowHeightStr,RowNumberStr,MoreLinkTypeStr,MoreLinkContentStr_R,CSSStyleStr,OpenTypeStr,DateCSSStyleStr,TxtNaviStr)
	Dim RsNewsObj,NewsSql,i,ClassSaveFilePath
	Dim TempDateShowStr,MoreLinkContentStr
	if Not AlreadyClassObj.Eof then
		OpenTypeStr = GetOpenTypeStr(OpenTypeStr)
		NewsSql = "Select top " & NewsListNumberStr & " NewsID,Title,Path,AddDate,HeadNewsTF,HeadNewsPath,FileName,FileExtName,TitleStyle from FS_News where ClassID='" & AlreadyClassObj("ClassID") & "' and AuditTF=1 and DelTF=0"
		GetOneClassNewsList = "<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""100%"">" & Chr(13) & Chr(10)
		NaviPicStr = GetNewsNavitionStr(TxtNaviStr,NaviPicStr)
		if RowHeightStr <> "" then RowHeightStr = " Height=""" & RowHeightStr & """"
		CompatPicStr = GetCompatPicStr(CompatPicStr,DateRightStr,DateRuleStr,RowNumberStr)
		if DateRuleStr <> "" then
			GetOneClassNewsList = GetOneClassNewsList & "<tr><td colspan=""2"" align=""center""><a " & OpenTypeStr & GetCSSStyleStr(CSSStyleStr) & " href=""" & GetOneClassLinkURL(AlreadyClassObj("ClassEName"),AlreadyClassObj("SaveFilePath"),AlreadyClassObj("FileExtName")) & """ >" & AlreadyClassObj("ClassCName") & "</a></td></tr>" & Chr(13) & Chr(10)
		else
			GetOneClassNewsList = GetOneClassNewsList & "<tr><td align=""center""><a " & OpenTypeStr & GetCSSStyleStr(CSSStyleStr) & " href=""" & GetOneClassLinkURL(AlreadyClassObj("ClassEName"),AlreadyClassObj("SaveFilePath"),AlreadyClassObj("FileExtName")) & """ >" & AlreadyClassObj("ClassCName") & "</a></td></tr>" & Chr(13) & Chr(10)
		end if
		Set RsNewsObj = Conn.Execute(NewsSql)
		if Not RsNewsObj.Eof then
			do while Not RsNewsObj.Eof
				GetOneClassNewsList = GetOneClassNewsList & "<tr>" & Chr(13) & Chr(10)
				for i = 1 to RowNumberStr
					if DateRuleStr <> "" then
						if DateRightStr = "Left" then
							TempDateShowStr = "&nbsp;&nbsp;<span " & GetCSSStyleStr(DateCSSStyleStr) & ">" & DateFormat(RsNewsObj("AddDate"),DateRuleStr) & "</span>"
						elseif DateRightStr = "Center" then
							TempDateShowStr = "<td align=""center""><span " & GetCSSStyleStr(DateCSSStyleStr) & ">" & DateFormat(RsNewsObj("AddDate"),DateRuleStr) & "</span>" & "</td>"& Chr(13) & Chr(10)
						elseif DateRightStr = "Right" then
							TempDateShowStr = "<td align=""Right""><span " & GetCSSStyleStr(DateCSSStyleStr) & ">" & DateFormat(RsNewsObj("AddDate"),DateRuleStr) & "</span>" & "</td>" & Chr(13) & Chr(10)
						else
							TempDateShowStr = "&nbsp;&nbsp;<span " & GetCSSStyleStr(DateCSSStyleStr) & ">" & DateFormat(RsNewsObj("AddDate"),DateRuleStr) & "</span>"
						end if
					else
						TempDateShowStr = ""
					end if
					if DateRightStr = "Center" OR DateRightStr = "Right" then
						GetOneClassNewsList = GetOneClassNewsList & "<td " & RowHeightStr & ">" & NaviPicStr & "<a " & OpenTypeStr & GetCSSStyleStr(CSSStyleStr) & " href=""" & GetOneNewsLinkURL(RsNewsObj("NewsID")) & """  title="""& RsNewsObj("Title")&""">" & GetHTMLTitle(RsNewsObj("TitleStyle"),GotTopic(RsNewsObj("Title"),TitleNumberStr)) & "</a></td>" & TempDateShowStr & Chr(13) & Chr(10)
					else
						GetOneClassNewsList = GetOneClassNewsList & "<td " & RowHeightStr & ">" & NaviPicStr & "<a " & OpenTypeStr & GetCSSStyleStr(CSSStyleStr) & " href=""" & GetOneNewsLinkURL(RsNewsObj("NewsID")) & """  title="""& RsNewsObj("Title")&""">" & GetHTMLTitle(RsNewsObj("TitleStyle"),GotTopic(RsNewsObj("Title"),TitleNumberStr)) & "</a>" & TempDateShowStr & "</td>" & Chr(13) & Chr(10)
					end if
					RsNewsObj.MoveNext
					if RsNewsObj.Eof then
						Exit For
					end if
				Next
				GetOneClassNewsList = GetOneClassNewsList & "</tr>" & Chr(13) & Chr(10) & CompatPicStr & Chr(13) & Chr(10)
			Loop
			if MoreLinkContentStr_R <> "" then
				if AlreadyClassObj("SaveFilePath") = "/" then
					ClassSaveFilePath = AlreadyClassObj("SaveFilePath")
				else
					ClassSaveFilePath = AlreadyClassObj("SaveFilePath") & "/"
				end if
				if MoreLinkTypeStr = "1" then
					MoreLinkContentStr="<a " & OpenTypeStr & GetCSSStyleStr(CSSStyleStr) & " href=""" & AvailableDoMain & ClassSaveFilePath & AlreadyClassObj("ClassEName") & "/index." & AlreadyClassObj("FileExtName") & """ ><img border=0 src=""" & MoreLinkContentStr_R & """></a>"
				elseif MoreLinkTypeStr = "0" then
					MoreLinkContentStr = "<a " & OpenTypeStr & GetCSSStyleStr(CSSStyleStr) & " href=""" & AvailableDoMain & ClassSaveFilePath & AlreadyClassObj("ClassEName") & "/index." & AlreadyClassObj("FileExtName") & """ >" & MoreLinkContentStr_R & "</a>"
				else
					MoreLinkContentStr = ""
				end if
				if DateRuleStr <> "" then
					GetOneClassNewsList = GetOneClassNewsList & "<tr><td " & GetRowSpanNumber(DateRightStr,DateRightStr,RowNumberStr) & " align=""right"">" & MoreLinkContentStr & "</td></tr>" & Chr(13) & Chr(10)
				else
					GetOneClassNewsList = GetOneClassNewsList & "<tr><td align=""right"" " & GetRowSpanNumber(DateRightStr,DateRightStr,RowNumberStr) & ">" & MoreLinkContentStr & "</td></tr>" & Chr(13) & Chr(10)
				end if
			end if
			GetOneClassNewsList = GetOneClassNewsList & "</table>" & Chr(13) & Chr(10)
		else
			GetOneClassNewsList = ""
		end if
		Set RsNewsObj = Nothing
	else
		GetOneClassNewsList = ""
	end if
End Function
'调用专题新闻列表
Function SpecialNewsList(SpecialID,NewsNumberStr,TitleNumberStr,CompatPicStr,NaviPicStr,DateRuleStr,DateRightStr,RowHeightStr,RowNumberStr,MoreLinkTypeStr,MoreLinkContentStr,CSSStyleStr,OpenTypeStr,DateCSSStyleStr,TxtNaviStr)
	Dim SpecialSql,RsSpecialObj,i,RsTempObj,ClassSaveFilePath
	Dim TempRowNumberStr,TempDateShowStr
	NaviPicStr = GetNewsNavitionStr(TxtNaviStr,NaviPicStr)
	if RowHeightStr <> "" then RowHeightStr = " Height=""" & RowHeightStr & """"
	TempRowNumberStr = GetRowSpanNumber(DateRuleStr,DateRightStr,RowNumberStr)
	CompatPicStr = GetCompatPicStr(CompatPicStr,DateRightStr,DateRuleStr,RowNumberStr)
	OpenTypeStr = GetOpenTypeStr(OpenTypeStr)
	SpecialSql = "Select * from FS_Special where EName ='" & SpecialID & "'"
	Set RsTempObj = Conn.Execute(SpecialSql)
	if Not RsTempObj.Eof then
		if RsTempObj("SaveFilePath") = "/" then
			ClassSaveFilePath = RsTempObj("SaveFilePath")
		else
			ClassSaveFilePath = RsTempObj("SaveFilePath") & "/"
		end if
		if MoreLinkContentStr <> "" then
			if MoreLinkTypeStr = "1" then
				MoreLinkContentStr="<a " & OpenTypeStr & GetCSSStyleStr(CSSStyleStr) & " href=""" & AvailableDoMain & ClassSaveFilePath & RsTempObj("EName") & "/index." & RsTempObj("FileExtName") & """ ><img border=0 src=""" & MoreLinkContentStr & """></a>"
			elseif MoreLinkTypeStr = "0" then
				MoreLinkContentStr = "<a " & OpenTypeStr & GetCSSStyleStr(CSSStyleStr) & " href=""" & AvailableDoMain & ClassSaveFilePath & RsTempObj("EName") & "/index." & RsTempObj("FileExtName") & """ >" & MoreLinkContentStr & "</a>"
			else
				MoreLinkContentStr = ""
			end if
		end if
		SpecialSql = "Select Top " & NewsNumberStr & " *,FS_NewsClass.FileExtName as ClassFileExtName,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.SpecialID like '%" & RsTempObj("SpecialID") & "%' and FS_News.AuditTF=1 and FS_News.DelTF=0 order by FS_News.ID Desc"
		Set RsSpecialObj = Conn.Execute(SpecialSql)
		SpecialNewsList = "<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""100%"">"
		do while Not RsSpecialObj.Eof
			SpecialNewsList = SpecialNewsList & "<tr>" & Chr(13) & Chr(10)
			for i = 1 to RowNumberStr
				if DateRuleStr <> "" then
					if DateRightStr = "Left" then
						TempDateShowStr = "&nbsp;&nbsp;<span " & GetCSSStyleStr(DateCSSStyleStr) & ">" & DateFormat(RsSpecialObj("AddDate"),DateRuleStr) & "</span>"
					elseif DateRightStr = "Center" then
						TempDateShowStr = "<td align=""center""><span " & GetCSSStyleStr(DateCSSStyleStr) & ">" & DateFormat(RsSpecialObj("AddDate"),DateRuleStr) & "</span>" & "</td>"& Chr(13) & Chr(10)
					elseif DateRightStr = "Right" then
						TempDateShowStr = "<td align=""Right""><span " & GetCSSStyleStr(DateCSSStyleStr) & ">" & DateFormat(RsSpecialObj("AddDate"),DateRuleStr) & "</td>" & "</span>" & Chr(13) & Chr(10)
					else
						TempDateShowStr = "&nbsp;&nbsp;<span " & GetCSSStyleStr(DateCSSStyleStr) & ">" & DateFormat(RsSpecialObj("AddDate"),DateRuleStr) & "</span>"
					end if
				else
					TempDateShowStr = ""
				end if
				if DateRightStr = "Center" OR DateRightStr = "Right" then
					SpecialNewsList = SpecialNewsList & "<td " & RowHeightStr & ">" & NaviPicStr & "<a " & OpenTypeStr & GetCSSStyleStr(CSSStyleStr) & " href=""" & GetOneNewsLinkURL(RsSpecialObj("NewsID")) & """  title="""& RsSpecialObj("Title")&""">" & GetHTMLTitle(RsSpecialObj("TitleStyle"),GotTopic(RsSpecialObj("Title"),TitleNumberStr)) & "</a></td>" & TempDateShowStr & Chr(13) & Chr(10)
				else
					SpecialNewsList = SpecialNewsList & "<td " & RowHeightStr & ">" & NaviPicStr & "<a " & OpenTypeStr & GetCSSStyleStr(CSSStyleStr) & " href=""" & GetOneNewsLinkURL(RsSpecialObj("NewsID")) & """  title="""& RsSpecialObj("Title")&""">" & GetHTMLTitle(RsSpecialObj("TitleStyle"),GotTopic(RsSpecialObj("Title"),TitleNumberStr)) & "</a>" & TempDateShowStr & "</td>" & Chr(13) & Chr(10)
				end if
				RsSpecialObj.MoveNext
				if RsSpecialObj.Eof then Exit For 
			Next
			SpecialNewsList = SpecialNewsList & "</tr>" & Chr(13) & Chr(10)
			SpecialNewsList = SpecialNewsList & CompatPicStr & Chr(13) & Chr(10)
		Loop
		if MoreLinkContentStr <> "" then
			SpecialNewsList = SpecialNewsList & "<tr><td align=""right"" " & TempRowNumberStr & ">" & MoreLinkContentStr & "</td></tr>" & Chr(13) & Chr(10)
		end if
		SpecialNewsList = SpecialNewsList & "</table>"
		Set RsSpecialObj = Nothing
	else
		SpecialNewsList = ""
	end if
	Set RsTempObj = Nothing
End Function
'专题终极分类
Function SpecialLastNewsList(NewsNumberStr,RowNumberStr,NaviPicStr,BGPicStr,RowHeightStr,CssFileStr,OpenModeStr,DetachPageStr,TitleNumberStr,DateRuleStr,DateRightStr,DateCSSStyleStr,TxtNaviStr)
	Dim SpecialSql,RsSpecialObj,RsTempObj,i,ClassSaveFilePath
	Dim PageNum,PageIndex,LoopVar,TempSpecialNewsList,SpecialNewsPageStr,TempClassSaveFilePath,ClassSaveExtName,j
	Dim TempRowNumberStr,TempDateShowStr
	if RefreshType = "Special" then
		NaviPicStr = GetNewsNavitionStr(TxtNaviStr,NaviPicStr)
		if RowHeightStr <> "" then RowHeightStr = " Height=""" & RowHeightStr & """"
		TempRowNumberStr = GetRowSpanNumber(DateRuleStr,DateRightStr,RowNumberStr)
		BGPicStr = GetCompatPicStr(BGPicStr,DateRightStr,DateRuleStr,RowNumberStr)
		OpenModeStr = GetOpenTypeStr(OpenModeStr)
		SpecialSql = "Select * from FS_Special where SpecialID ='" & RefreshID & "'"
		Set RsTempObj = Conn.Execute(SpecialSql)
		if Not RsTempObj.Eof then
			if RsTempObj("SaveFilePath") = "/" then
				TempClassSaveFilePath = AvailableDoMain & RsTempObj("SaveFilePath") & RsTempObj("EName")
				ClassSaveExtName = RsTempObj("FileExtName")
			else
				TempClassSaveFilePath = AvailableDoMain & RsTempObj("SaveFilePath") & "/" & RsTempObj("EName")
				ClassSaveExtName = RsTempObj("FileExtName")
			end if
			SpecialSql = "Select *,FS_NewsClass.FileExtName as ClassFileExtName,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass where FS_News.ClassID=FS_NewsClass.ClassID  and FS_News.SpecialID like '%" & RefreshID & "%' and FS_News.AuditTF=1 and FS_News.DelTF=0 order by FS_News.ID Desc"
			Set RsSpecialObj = Server.CreateObject(G_FS_RS)
			RsSpecialObj.Open SpecialSql,Conn,1,1
			if Not RsSpecialObj.Eof then
				RsSpecialObj.PageSize = NewsNumberStr
				PageNum = RsSpecialObj.PageCount
			else
				PageNum = 0
			end if
			if (DetachPageStr = "1") and (PageNum > 1) then
				for PageIndex = 1 to PageNum
					SpecialNewsPageStr = "<tr><td " & TempRowNumberStr & "><table border=""0"" width=""100%""><tr><td width=""50%"" align=""right""><br>第<font color=red>" & PageIndex&"</font>页,共<font color=red>"&PageNum & "</font>页 "
					if PageIndex = 1 then
						SpecialNewsPageStr = SpecialNewsPageStr & "<font face=webdings>9</font> "
						SpecialNewsPageStr = SpecialNewsPageStr & "<font face=webdings>7</font> "
					elseif PageIndex=2 then
						SpecialNewsPageStr = SpecialNewsPageStr & "<a href=""" & TempClassSaveFilePath & "/" & "index." & ClassSaveExtName & """ title=首页><font face=webdings>9</font></a> "
						SpecialNewsPageStr = SpecialNewsPageStr & "<a href=""" & TempClassSaveFilePath & "/" & "index." & ClassSaveExtName & """ title=上一页><font face=webdings>7</font></a> "
					else
						SpecialNewsPageStr = SpecialNewsPageStr & "<a href=""" & TempClassSaveFilePath & "/" & "index." & ClassSaveExtName & """ title=首页><font face=webdings>9</font></a> "
						SpecialNewsPageStr = SpecialNewsPageStr & "<a href=""" & TempClassSaveFilePath & "/" & "index_"&PageIndex-1&"." & ClassSaveExtName & """　title=上一页><font face=webdings>7</font></a> "
					end if
					dim G
					G=0
					for j = PageIndex to PageNum
						if j = 1 then
							if j=PageIndex then
								SpecialNewsPageStr = SpecialNewsPageStr & "<a href=""" & TempClassSaveFilePath & "/" & "index" & "." & ClassSaveExtName & """><font color=red>[" & j & "]</font></a> "
							else
								SpecialNewsPageStr = SpecialNewsPageStr & "<a href=""" & TempClassSaveFilePath & "/" & "index" & "." & ClassSaveExtName & """>[" & j & "]</a> "
							end if
						else
							if j=PageIndex then
								SpecialNewsPageStr = SpecialNewsPageStr & "<a href=""" & TempClassSaveFilePath & "/" & "index" & "_" & j & "." & ClassSaveExtName & """><font color=red>[" & j & "]</font></a> "
							else
								SpecialNewsPageStr = SpecialNewsPageStr & "<a href=""" & TempClassSaveFilePath & "/" & "index" & "_" & j & "." & ClassSaveExtName & """>[" & j & "]</a> "
							end if
						end if
						G=G+1
						if G mod 10 = 0 then
					     exit for
						End if
					Next
					if PageIndex=PageNum then
						SpecialNewsPageStr = SpecialNewsPageStr & "<font face=webdings>8</font> "
					else
						SpecialNewsPageStr = SpecialNewsPageStr & "<a href=""" & TempClassSaveFilePath & "/" & "index" & "_" & PageIndex+1 & "." & ClassSaveExtName & """  title=下一页><font face=webdings>8</font></a> "
					end if
					if PageIndex=PageNum then
						SpecialNewsPageStr = SpecialNewsPageStr & "<font face=webdings>:</font> "
					else
						SpecialNewsPageStr = SpecialNewsPageStr & "<a href=""" & TempClassSaveFilePath & "/" & "index_"&PageNum&"." & ClassSaveExtName & """ title=最后一页><font face=webdings>:</font></a> "
					end if
					SpecialNewsPageStr = SpecialNewsPageStr & "</td></tr></table></td></tr>"
					RsSpecialObj.AbsolutePage = PageIndex
					Dim TempAlreadyShow
					TempAlreadyShow = 1
					TempSpecialNewsList = "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">" & Chr(13) & Chr(10)
					for LoopVar = 1 to RsSpecialObj.PageSize
						if TempAlreadyShow > RsSpecialObj.PageSize then Exit For
						if RsSpecialObj.Eof then Exit For
						TempSpecialNewsList = TempSpecialNewsList & "<tr>" & Chr(13) & Chr(10)
						for i = 1 to RowNumberStr
							TempAlreadyShow = TempAlreadyShow + 1
							if DateRuleStr <> "" then
								if DateRightStr = "Left" then
									TempDateShowStr = "&nbsp;&nbsp;<span " & GetCSSStyleStr(DateCSSStyleStr) & ">" & DateFormat(RsSpecialObj("AddDate"),DateRuleStr) & "</span>"
								elseif DateRightStr = "Center" then
									TempDateShowStr = "<td align=""center""><span " & GetCSSStyleStr(DateCSSStyleStr) & ">" & DateFormat(RsSpecialObj("AddDate"),DateRuleStr) & "</span>" & "</td>"& Chr(13) & Chr(10)
								elseif DateRightStr = "Right" then
									TempDateShowStr = "<td align=""Right""><span " & GetCSSStyleStr(DateCSSStyleStr) & ">" & DateFormat(RsSpecialObj("AddDate"),DateRuleStr) & "</span>" & "</td>"& Chr(13) & Chr(10)
								else
									TempDateShowStr = "&nbsp;&nbsp;<span " & GetCSSStyleStr(DateCSSStyleStr) & ">" & DateFormat(RsSpecialObj("AddDate"),DateRuleStr) & "</span>"
								end if
							else
								TempDateShowStr = ""
							end if
							if DateRightStr = "Center" OR DateRightStr = "Right" then
								TempSpecialNewsList = TempSpecialNewsList & "<td " & RowHeightStr & ">" & NaviPicStr & "<a " & OpenModeStr & GetCSSStyleStr(CssFileStr) & " href=""" & GetOneNewsLinkURL(RsSpecialObj("NewsID")) & """  title="""& RsSpecialObj("Title")&""">" & GetHTMLTitle(RsSpecialObj("TitleStyle"),GotTopic(RsSpecialObj("Title"),TitleNumberStr)) & "</a></td>" & TempDateShowStr & Chr(13) & Chr(10)
							else
								TempSpecialNewsList = TempSpecialNewsList & "<td " & RowHeightStr & ">" & NaviPicStr & "<a " & OpenModeStr & GetCSSStyleStr(CssFileStr) & " href=""" & GetOneNewsLinkURL(RsSpecialObj("NewsID")) & """  title="""& RsSpecialObj("Title")&""">" & GetHTMLTitle(RsSpecialObj("TitleStyle"),GotTopic(RsSpecialObj("Title"),TitleNumberStr)) & "</a>" & TempDateShowStr & "</td>" & Chr(13) & Chr(10)
							end if
							RsSpecialObj.MoveNext
							if RsSpecialObj.Eof then Exit For
							if TempAlreadyShow > RsSpecialObj.PageSize then Exit For
						Next
						TempSpecialNewsList = TempSpecialNewsList & "</tr>" & Chr(13) & Chr(10)
						TempSpecialNewsList = TempSpecialNewsList & BGPicStr & Chr(13) & Chr(10)
					Next
					TempSpecialNewsList = TempSpecialNewsList & SpecialNewsPageStr & Chr(13) & Chr(10)
					TempSpecialNewsList = TempSpecialNewsList & "</table>"
					if SpecialLastNewsList = "" then
						SpecialLastNewsList = TempSpecialNewsList
					else
						SpecialLastNewsList = SpecialLastNewsList & "$$$" & TempSpecialNewsList
					end if
				Next
			else
				SpecialLastNewsList = "<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""100%"">"
				do while Not RsSpecialObj.Eof
					SpecialLastNewsList = SpecialLastNewsList & "<tr>" & Chr(13) & Chr(10)
					for i = 1 to RowNumberStr
						if DateRuleStr <> "" then
							if DateRightStr = "Left" then
								TempDateShowStr = "&nbsp;&nbsp;<span " & GetCSSStyleStr(DateCSSStyleStr) & ">" & DateFormat(RsSpecialObj("AddDate"),DateRuleStr) & "</span>"
							elseif DateRightStr = "Center" then
								TempDateShowStr = "<td align=""center""><span " & GetCSSStyleStr(DateCSSStyleStr) & ">" & DateFormat(RsSpecialObj("AddDate"),DateRuleStr) & "</span>" & "</td>"& Chr(13) & Chr(10)
							elseif DateRightStr = "Right" then
								TempDateShowStr = "<td align=""Right""><span " & GetCSSStyleStr(DateCSSStyleStr) & ">" & DateFormat(RsSpecialObj("AddDate"),DateRuleStr)& "</span>" & "</td>"  & Chr(13) & Chr(10)
							else
								TempDateShowStr = "&nbsp;&nbsp;<span " & GetCSSStyleStr(DateCSSStyleStr) & ">" & DateFormat(RsSpecialObj("AddDate"),DateRuleStr) & "</span>"
							end if
						else
							TempDateShowStr = ""
						end if
						if DateRightStr = "Center" OR DateRightStr = "Right" then
							SpecialLastNewsList = SpecialLastNewsList & "<td " & RowHeightStr & ">" & NaviPicStr & "<a " & OpenModeStr & GetCSSStyleStr(CssFileStr) & " href=""" & GetOneNewsLinkURL(RsSpecialObj("NewsID")) & """  title="""& RsSpecialObj("Title")&""">" & GetHTMLTitle(RsSpecialObj("TitleStyle"),GotTopic(RsSpecialObj("Title"),TitleNumberStr)) & "</a></td>" & TempDateShowStr & Chr(13) & Chr(10)
						else
							SpecialLastNewsList = SpecialLastNewsList & "<td " & RowHeightStr & ">" & NaviPicStr & "<a " & OpenModeStr & GetCSSStyleStr(CssFileStr) & " href=""" & GetOneNewsLinkURL(RsSpecialObj("NewsID")) & """  title="""& RsSpecialObj("Title")&""">" & GetHTMLTitle(RsSpecialObj("TitleStyle"),GotTopic(RsSpecialObj("Title"),TitleNumberStr)) & "</a>" & TempDateShowStr & "</td>" & Chr(13) & Chr(10)
						end if
						RsSpecialObj.MoveNext
						if RsSpecialObj.Eof then Exit For
					Next
					SpecialLastNewsList = SpecialLastNewsList & "</tr>"
					SpecialLastNewsList = SpecialLastNewsList & BGPicStr
				Loop
				SpecialLastNewsList = SpecialLastNewsList & "</table>"
			end if
			Set RsSpecialObj = Nothing
			SpecialLastNewsList = Split(SpecialLastNewsList,"$$$")
		else
			SpecialLastNewsList = Array("")
		end if
	else
		SpecialLastNewsList = Array("")
	end if
End Function
'专题图片导航
Function SpecialNavi(PicTF,SpecialClassID,PicHeight,PicWidth,Dhang,SpecialCss,SpecialMore)
	Dim SpecialNaviObj,ClassSaveFilePath
	Set SpecialNaviObj=conn.execute("select * from FS_special where EName='"&SpecialClassID&"'")
	if SpecialNaviObj.eof then
		SpecialNavi=""
	else
		if SpecialNaviObj("SaveFilePath") = "/" then
			ClassSaveFilePath = SpecialNaviObj("SaveFilePath")
		else
			ClassSaveFilePath = SpecialNaviObj("SaveFilePath") & "/"
		end if
		SpecialNavi=""
		if PicTF = "1" then 
			SpecialNavi=SpecialNavi&"<table><tr><td><img src="  & AvailableDoMain & SpecialNaviObj("NaviPic") & " height=" & PicHeight & " width=" & PicWidth & "></td><td><a href=" & AvailableDoMain & ClassSaveFilePath & SpecialNaviObj("EName") & "/index." & SpecialNaviObj("FileExtName") & " target=_blank><b>"&SpecialNaviObj("CName")&"</b></a><br>&nbsp;&nbsp;&nbsp;&nbsp;<span class="&SpecialCss&">"&left(SpecialNaviObj("IndexNaviWord"),Dhang)&"</span></td></tr>"
		else
			SpecialNavi=SpecialNavi&"<table><tr><td><a href=" & AvailableDoMain & ClassSaveFilePath & SpecialNaviObj("EName") & "/index." & SpecialNaviObj("FileExtName") & " target=_blank><b>" & SpecialNaviObj("CName") & "</b></a><br>&nbsp;&nbsp;&nbsp;&nbsp;<span class=" & SpecialCss & ">" & left(SpecialNaviObj("IndexNaviWord"),Dhang)&"</span></td></tr>"
		end if
			SpecialNavi=SpecialNavi&"</table>"
	end if
	Set SpecialNaviObj = Nothing
	
End Function
'当前位置
Function Location(NaviType,CompatStr,OpenTypeStr,CSSStyleStr)
	OpenTypeStr = GetOpenTypeStr(OpenTypeStr)
	Select Case RefreshType
		Case "Index"
			Location = "<a " & GetCSSStyleStr(CSSStyleStr) & " href=""" & AvailableDoMain & "/Index."&confimsn("IndexExtName")&"""><font color=red>首页</font></a>"
		Case "Class"
			Location = GetClassLocationStr(RefreshID,NaviType,CompatStr,OpenTypeStr,CSSStyleStr)
		Case "News"
			Location = GetNewsLocationStr(RefreshID,NaviType,CompatStr,OpenTypeStr,CSSStyleStr)
		Case "Special"
			Location = GetSpecialLocationStr(RefreshID,NaviType,CompatStr,OpenTypeStr,CSSStyleStr)
		Case "DownLoad"
			Location = GetDownLoadLocationStr(RefreshID,NaviType,CompatStr,OpenTypeStr,CSSStyleStr)
		Case Else
			Location = ""
	End Select
End Function
'栏目当前位置
Function GetClassLocationStr(ClassID,NaviType,CompatStr,OpenTypeStr,CSSStyleStr)
	Dim SqlClass,RsClassObj
	if ClassID = "" then Exit Function
	if NaviType = "1" then
		CompatStr = "<img src=""" & AvailableDoMain & CompatStr & """>"
	end if
	Set RsClassObj = Conn.Execute("Select FileExtName,SaveFilePath,ParentID,ClassID,ClassEName,ClassCName from FS_NewsClass where ClassID='" & ClassID & "'")
	if Not RsClassObj.Eof then
		GetClassLocationStr = GetClassLocationStr & "<a " & OpenTypeStr & GetCSSStyleStr(CSSStyleStr) & " href=""" & GetOneClassLinkURL(RsClassObj("ClassEName"),RsClassObj("SaveFilePath"),RsClassObj("FileExtName")) & """ >" & RsClassObj("ClassCName") & "</a>"
		do while RsClassObj("ParentID") <> 0
			Set RsClassObj = Conn.Execute("Select FileExtName,SaveFilePath,ParentID,ClassID,ClassEName,ClassCName from FS_NewsClass where ClassID='" & RsClassObj("ParentID") & "'")
			if RsClassObj.Eof then Exit do
			GetClassLocationStr = "<a " & OpenTypeStr & GetCSSStyleStr(CSSStyleStr) & " href=""" & GetOneClassLinkURL(RsClassObj("ClassEName"),RsClassObj("SaveFilePath"),RsClassObj("FileExtName")) & """ >" & RsClassObj("ClassCName") & "</a>" & CompatStr & GetClassLocationStr
		loop
	end if
	GetClassLocationStr = "<a " & OpenTypeStr & GetCSSStyleStr(CSSStyleStr) & " href=""" & AvailableDoMain & "/Index." & Confimsn("IndexExtName") & """><font color=red>首页</font></a>" & CompatStr & GetClassLocationStr
	RsClassObj.Close
	Set RsClassObj = Nothing
End Function
'新闻当前位置
Function GetNewsLocationStr(NewsID,NaviType,CompatStr,OpenTypeStr,CSSStyleStr)
	Dim RsNewsObj
	if NaviType = "1" then
		CompatStr = "<img src=""" & AvailableDoMain & CompatStr & """>"
	end if
	Set RsNewsObj = Conn.Execute("Select ClassID from FS_News where NewsID='" & NewsID & "'")
	if Not RsNewsObj.Eof then
		GetNewsLocationStr = GetClassLocationStr(RsNewsObj("ClassID"),NaviType,CompatStr,OpenTypeStr,CSSStyleStr) & CompatStr & "正文"
	else
		GetNewsLocationStr = ""
	end if
	Set RsNewsObj = Nothing
End Function
'下载当前位置
Function GetDownLoadLocationStr(DownLoadID,NaviType,CompatStr,OpenTypeStr,CSSStyleStr)
	Dim RsDownLoadObj
	if NaviType = "1" then
		CompatStr = "<img src=""" & AvailableDoMain & CompatStr & """>"
	end if
	Set RsDownLoadObj = Conn.Execute("Select ClassID from FS_DownLoad where DownLoadID='" & DownLoadID & "'")
	if Not RsDownLoadObj.Eof then
		GetDownLoadLocationStr = GetClassLocationStr(RsDownLoadObj("ClassID"),NaviType,CompatStr,OpenTypeStr,CSSStyleStr) & CompatStr & "下载"
	else
		GetDownLoadLocationStr = ""
	end if
	Set RsDownLoadObj = Nothing
End Function
'专题当前位置
Function GetSpecialLocationStr(SpecialID,NaviType,CompatStr,OpenTypeStr,CSSStyleStr)
	Dim SpecialSql,RsSpecialObj
	if NaviType = "1" then
		CompatStr = "<img src=""" & AvailableDoMain & CompatStr & """>"
	end if
	SpecialSql = "Select * from FS_Special where SpecialID='" & SpecialID & "'"
	Set RsSpecialObj = Conn.Execute(SpecialSql)
	if RsSpecialObj.Eof then
		GetSpecialLocationStr = ""
	else
		GetSpecialLocationStr = "<a " & OpenTypeStr & GetCSSStyleStr(CSSStyleStr) & " href=""" & AvailableDoMain & "/Index."&confimsn("IndexExtName")&"""><font color=red>首页</font></a>" & CompatStr & RsSpecialObj("CName") & "专题"
	end if
	Set RsSpecialObj = Nothing
End Function
'总站导航
Function LocationNavi(NaviType,RowNumber,NaviPicStr,CompatPicStr,OpenTypeStr,CSSStyleStr,TxtNaviStr)
	Dim NaviArray,i,TempStr
	NaviPicStr = GetNewsNavitionStr(TxtNaviStr,NaviPicStr)
	if CompatPicStr <> "" then
		CompatPicStr = "background=""" & CompatPicStr & """"
	else
		CompatPicStr = ""
	end if
	RowNumber = RowNumber
	LocationNavi = "<table border=""0"" width=""100%;"">" & Chr(13) & Chr(10) & "<tr>" & Chr(13) & Chr(10)
	Select Case NaviType
		Case "1"
			TempStr = GetRootClassNavi(OpenTypeStr,CSSStyleStr)
		Case "2"
			TempStr = GetSpecialNavi(OpenTypeStr,CSSStyleStr)
		Case "3"
			TempStr = GetPlusNavi(OpenTypeStr,CSSStyleStr)
		Case "4"
			TempStr = GetRootClassNavi(OpenTypeStr,CSSStyleStr) & "{$$$}" & GetSpecialNavi(OpenTypeStr,CSSStyleStr)
		Case "5"
			TempStr = GetRootClassNavi(OpenTypeStr,CSSStyleStr) & "{$$$}" & GetPlusNavi(OpenTypeStr,CSSStyleStr)
		Case "6"
			TempStr = GetSpecialNavi(OpenTypeStr,CSSStyleStr) & "{$$$}" & GetPlusNavi(OpenTypeStr,CSSStyleStr)
		Case "7"
			TempStr = GetRootClassNavi(OpenTypeStr,CSSStyleStr) & "{$$$}" & GetSpecialNavi(OpenTypeStr,CSSStyleStr) & "{$$$}" & GetPlusNavi(OpenTypeStr,CSSStyleStr)
		Case Else
			LocationNavi = ""
			Exit Function
	End Select
	If Left(Trim(TempStr),5)="{$$$}" then TempStr=Mid(TempStr,6)
	NaviArray = Split(TempStr,"{$$$}")
	for i = LBound(NaviArray) to UBound(NaviArray)
		LocationNavi = LocationNavi & "<td " & CompatPicStr & ">" & NaviPicStr & NaviArray(i) & "</td>" & Chr(13) & Chr(10)
		if ((i + 1) Mod RowNumber) = 0 then
			LocationNavi = LocationNavi & "</tr>" & Chr(13) & Chr(10) & "<tr>"
		end if
	Next
	LocationNavi = LocationNavi & "</tr>" & Chr(13) & Chr(10) & "</table>" & Chr(13) & Chr(10)
End Function
'专题导航
Function GetSpecialNavi(OpenTypeStr,CSSStyleStr)
	Dim SpecialSql,RsSpecialObj
	OpenTypeStr = GetOpenTypeStr(OpenTypeStr)
	SpecialSql = "Select * from FS_Special where ShowNaviTF=1 order by ID desc"
	Set RsSpecialObj = Conn.Execute(SpecialSql)
	do while Not RsSpecialObj.Eof
		if GetSpecialNavi = "" then
			if RsSpecialObj("SaveFilePath") = "/" then
				GetSpecialNavi = "<a " & GetCSSStyleStr(CSSStyleStr) & OpenTypeStr & " href=""" & AvailableDoMain & RsSpecialObj("SaveFilePath") & RsSpecialObj("EName") & "/" & "index." & RsSpecialObj("FileExtName") & """ >" & RsSpecialObj("CName") & "</a>"
			else
				GetSpecialNavi = "<a " & GetCSSStyleStr(CSSStyleStr) & OpenTypeStr & " href=""" & AvailableDoMain & RsSpecialObj("SaveFilePath") & "/" & RsSpecialObj("EName") & "/" & "index." & RsSpecialObj("FileExtName") & """ >" & RsSpecialObj("CName") & "</a>"
			end if
		else
			if RsSpecialObj("SaveFilePath") = "/" then
				GetSpecialNavi = GetSpecialNavi & "{$$$}" & "<a " & GetCSSStyleStr(CSSStyleStr) & OpenTypeStr & " href=""" & AvailableDoMain & RsSpecialObj("SaveFilePath") & RsSpecialObj("EName") & "/" & "index." & RsSpecialObj("FileExtName") & """ >" & RsSpecialObj("CName") & "</a>"
			else
				GetSpecialNavi = GetSpecialNavi & "{$$$}" & "<a " & GetCSSStyleStr(CSSStyleStr) & OpenTypeStr & " href=""" & AvailableDoMain & RsSpecialObj("SaveFilePath") & "/" & RsSpecialObj("EName") & "/" & "index." & RsSpecialObj("FileExtName") & """ >" & RsSpecialObj("CName") & "</a>"
			end if
		end if
		RsSpecialObj.MoveNext
	Loop
End Function
'栏目导航
Function GetRootClassNavi(OpenTypeStr,CSSStyleStr)
	Dim ClassSql,RsClassObj
	OpenTypeStr = GetOpenTypeStr(OpenTypeStr)
	ClassSql = "Select IsOutClass,SaveFilePath,ClassEName,ClassCName,FileExtName,ShowTF,DoMain,ClassLink from FS_NewsClass where ParentID='0' and DelFlag=0 and ShowTF=1 order by orders desc"
	Set RsClassObj = Conn.Execute(ClassSql)
	do while Not RsClassObj.Eof
		if RsClassObj("IsOutClass") = "1" then
			GetRootClassNavi = GetRootClassNavi & "{$$$}" &  "<a " & GetCSSStyleStr(CSSStyleStr) & OpenTypeStr & " href=""" & RsClassObj("ClassLink") & """ >" & RsClassObj("ClassCName") & "</a>"
		else
			GetRootClassNavi = GetRootClassNavi & "{$$$}" & "<a " & GetCSSStyleStr(CSSStyleStr) & OpenTypeStr & " href=""" & GetOneClassLinkURL(RsClassObj("ClassEName"),RsClassObj("SaveFilePath"),RsClassObj("FileExtName")) & """ >" & RsClassObj("ClassCName") & "</a>"
		end if
		RsClassObj.MoveNext
	loop
	Set RsClassObj = Nothing
End Function
'插件导航
Function GetPlusNavi(OpenTypeStr,CSSStyleStr)
	Dim PlusSql,RsPlusObj,OpenType
	PlusSql = "Select Name,Link,OpenType from FS_Plus where ShowTF=1  order by ID asc"
	Set RsPlusObj = Conn.Execute(PlusSql)
	do while Not RsPlusObj.Eof
		if RsPlusObj("OpenType") = 1 then
			OpenType = " target=""_blank"""
		else
			OpenType = ""
		end if
		if GetPlusNavi = "" then
			GetPlusNavi = "<a " & GetCSSStyleStr(CSSStyleStr) & OpenType & " href=""" & RsPlusObj("Link") & """ >" & RsPlusObj("Name") & "</a>"
		else
			GetPlusNavi = GetPlusNavi & "{$$$}" & "<a " & GetCSSStyleStr(CSSStyleStr) & OpenType & " href=""" & RsPlusObj("Link") & """ >" & RsPlusObj("Name") & "</a>"
		end if
		RsPlusObj.MoveNext
	loop
	Set RsPlusObj = Nothing
End Function
'栏目导航
Function ClassNavi(NaviPicStr,CompatPicStr,RowNumberStr,OpenTypeStr,CSSStyleStr,TxtNaviStr)
	Dim ClassSql,RsClassObj,i
	if RefreshType = "Class" then
		CompatPicStr = GetCompatPicStr(CompatPicStr,"","",RowNumberStr)
		OpenTypeStr = GetOpenTypeStr(OpenTypeStr)
		NaviPicStr = GetNewsNavitionStr(TxtNaviStr,NaviPicStr)
		ClassSql = "Select SaveFilePath,ClassEName,ClassCName,FileExtName from FS_NewsClass where ShowTF=1 and DelFlag=0 and ParentID='" & RefreshID & "' order by orders desc"
		Set RsClassObj = Conn.Execute(ClassSql)
		ClassNavi = "<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""100%"">" & Chr(13) & Chr(10)
		do while Not RsClassObj.Eof
			ClassNavi = ClassNavi & "<tr>" & Chr(13) & Chr(10)
			for i = 1 to RowNumberStr
				ClassNavi = ClassNavi & "<td>" & NaviPicStr & "<a " & GetCSSStyleStr(CSSStyleStr) & OpenTypeStr & " href=""" & GetOneClassLinkURL(RsClassObj("ClassEName"),RsClassObj("SaveFilePath"),RsClassObj("FileExtName")) & """ >" & RsClassObj("ClassCName") & "</a></td>" & Chr(13) & Chr(10)
				RsClassObj.MoveNext
				if RsClassObj.Eof then Exit For
			Next
			ClassNavi = ClassNavi & "</tr>" & Chr(13) & Chr(10)
		loop
		ClassNavi = ClassNavi & "</table>" & Chr(13) & Chr(10)
		Set RsClassObj = Nothing
	else
		ClassNavi = ""
	end if
End Function
'热点新闻
Function HotNews(ClassEName,SoonClassStr,NewNumberStr,TitleNumberStr,RowNumberStr,NaviPicStr,CompatPicStr,OpenTypeStr,CSSStyleStr,RowHeightStr,TxtNaviStr)
	Dim HotNewsSql,RsHotNewsObj,i
	TitleNumberStr = GetTitleNumberStr(TitleNumberStr)
	NewNumberStr = GetTitleNumberStr(NewNumberStr)
	OpenTypeStr = GetOpenTypeStr(OpenTypeStr)
	CompatPicStr = GetCompatPicStr(CompatPicStr,"","",RowNumberStr)
	NaviPicStr = GetNewsNavitionStr(TxtNaviStr,NaviPicStr)
	if RowHeightStr <> "" then RowHeightStr = " Height=""" & RowHeightStr & """"
	'----------------------
	dim TemppID,TemppSql,EndClassIDList
	If ClassEName<>"" then
		If SoonClassStr="1" then 
			TemppSql="select ClassID from FS_NewsClass where ClassEName='" & ClassEName & "'"
			Set TemppID=conn.execute(TemppSql)
			EndClassIDList= "'" & TemppID(0) & "'" & AllChildClassIDStrList(TemppID(0))
		Else
			TemppSql="select ClassID from FS_NewsClass where ClassEName='" & ClassEName & "'"
			Set TemppID=conn.execute(TemppSql)
			EndClassIDList="'" & TemppID(0) & "'"
		End if
	Else
		EndClassIDList=""
	end if

	if EndClassIDList <> "" then
		HotNewsSql = "Select top "&NewNumberStr&" *,FS_NewsClass.FileExtName as ClassFileExtName,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass where FS_News.ClassID=FS_NewsClass.ClassID and DelTF=0 and FS_News.AuditTF=1 and FS_News.ClassID in (" & EndClassIDList & ") order by FS_News.ClickNum Desc,FS_News.id desc" 
	else
		HotNewsSql = "Select top "&NewNumberStr&" *,FS_NewsClass.FileExtName as ClassFileExtName,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass where FS_News.ClassID=FS_NewsClass.ClassID and DelTF=0 and FS_News.AuditTF=1 order by FS_News.ClickNum Desc,FS_News.id desc" 
	end if
	Set RsHotNewsObj = Conn.Execute(HotNewsSql)
	HotNews = "<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""100%"">" & Chr(13) & Chr(10)
	do while Not RsHotNewsObj.Eof
		HotNews = HotNews & "<tr>" & Chr(13) & Chr(10)
		for i = 1 to RowNumberStr
			HotNews = HotNews & "<td " & RowHeightStr & ">" & NaviPicStr & "<a " & GetCSSStyleStr(CSSStyleStr) & OpenTypeStr & " href=""" & GetOneNewsLinkURL(RsHotNewsObj("NewsID")) & """  title="""& RsHotNewsObj("Title")&""">" & GetHTMLTitle(RsHotNewsObj("TitleStyle"),GotTopic(RsHotNewsObj("Title"),TitleNumberStr)) & "</a></td>" & Chr(13) & Chr(10)
			RsHotNewsObj.MoveNext
			if RsHotNewsObj.Eof then Exit For
		Next
		HotNews = HotNews & "</tr>" & Chr(13) & Chr(10) & CompatPicStr & Chr(13) & Chr(10)
	loop
	Set RsHotNewsObj = Nothing
	HotNews = HotNews & "</table>"
End Function 
'最新新闻 
Function LastNews(ClassEName,SoonClassStr,NewNumberStr,TitleNumberStr,RowNumberStr,NaviPicStr,CompatPicStr,OpenTypeStr,CSSStyleStr,RowHeightStr,TxtNaviStr) 
	Dim LastNewsSql,RsLastNewsObj,i
	TitleNumberStr = GetTitleNumberStr(TitleNumberStr)
	OpenTypeStr = GetOpenTypeStr(OpenTypeStr)
	CompatPicStr = GetCompatPicStr(CompatPicStr,"","",RowNumberStr)
	NaviPicStr = GetNewsNavitionStr(TxtNaviStr,NaviPicStr)
	if RowHeightStr <> "" then RowHeightStr = " Height=""" & RowHeightStr & """"
	'---------
	dim TemppID,TemppSql,EndClassIDList
	If ClassEName<>"" then
		If SoonClassStr="1" then 
			TemppSql="select ClassID from FS_NewsClass where ClassEName='" & ClassEName & "'"
			Set TemppID=conn.execute(TemppSql)
			EndClassIDList= "'" & TemppID(0) & "'" & AllChildClassIDStrList(TemppID(0))
		Else
			TemppSql="select ClassID from FS_NewsClass where ClassEName='" & ClassEName & "'"
			Set TemppID=conn.execute(TemppSql)
			EndClassIDList="'" & TemppID(0) & "'"
		End if
	Else
		EndClassIDList=""
	end if	
	if EndClassIDList <> "" then	
		LastNewsSql = "Select Top " & NewNumberStr & " *,FS_NewsClass.FileExtName as ClassFileExtName,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass where FS_News.ClassID=FS_NewsClass.ClassID and DelTF=0 and FS_News.AuditTF=1 and FS_News.Classid in(" & EndClassIDList & ") order by FS_News.ID Desc"
	else
		LastNewsSql = "Select Top " & NewNumberStr & " *,FS_NewsClass.FileExtName as ClassFileExtName,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass where FS_News.ClassID=FS_NewsClass.ClassID and DelTF=0 and FS_News.AuditTF=1 order by FS_News.ID Desc"
	end if
	Set RsLastNewsObj = Conn.Execute(LastNewsSql)
	LastNews = "<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""100%"">" & Chr(13) & Chr(10)
	do while Not RsLastNewsObj.Eof
		LastNews = LastNews & "<tr>" & Chr(13) & Chr(10)
		for i = 1 to RowNumberStr
			LastNews = LastNews & "<td " & RowHeightStr & ">" & NaviPicStr & "<a " & GetCSSStyleStr(CSSStyleStr) & OpenTypeStr & " href=""" & GetOneNewsLinkURL(RsLastNewsObj("NewsID")) & """ title="""& RsLastNewsObj("Title")&""">" & GetHTMLTitle(RsLastNewsObj("TitleStyle"),GotTopic(RsLastNewsObj("Title"),TitleNumberStr)) & "</a></td>" & Chr(13) & Chr(10)
			RsLastNewsObj.MoveNext
			if RsLastNewsObj.Eof then
				Exit For
			end if
		Next
		LastNews = LastNews & "</tr>" & Chr(13) & Chr(10) & CompatPicStr & Chr(13) & Chr(10)
	loop
	Set RsLastNewsObj = Nothing
	LastNews = LastNews & "</table>"
End Function
'推荐新闻
Function RecNews(ClassEName,SoonClassStr,NewNumberStr,TitleNumberStr,RowNumberStr,NaviPicStr,CompatPicStr,OpenTypeStr,CSSStyleStr,RowHeightStr,TxtNaviStr)
	Dim RecNewsSql,RsRecNewsObj,i
	TitleNumberStr = GetTitleNumberStr(TitleNumberStr)
	OpenTypeStr = GetOpenTypeStr(OpenTypeStr)
	CompatPicStr = GetCompatPicStr(CompatPicStr,"","",RowNumberStr)
	NaviPicStr = GetNewsNavitionStr(TxtNaviStr,NaviPicStr)
	if RowHeightStr <> "" then RowHeightStr = " Height=""" & RowHeightStr & """"
	'-------------
	dim TemppID,TemppSql,EndClassIDList
	If ClassEName<>"" then
		If SoonClassStr="1" then 
			TemppSql="select ClassID from FS_NewsClass where ClassEName='" & ClassEName & "'"
			Set TemppID=conn.execute(TemppSql)
			EndClassIDList= "'" & TemppID(0) & "'" & AllChildClassIDStrList(TemppID(0))
		Else
			TemppSql="select ClassID from FS_NewsClass where ClassEName='" & ClassEName & "'"
			Set TemppID=conn.execute(TemppSql)
			EndClassIDList="'" & TemppID(0) & "'"
		End if
	Else
		EndClassIDList=""
	end if


	if EndClassIDList <> "" then
		RecNewsSql = "Select Top " & NewNumberStr & " *,FS_NewsClass.FileExtName as ClassFileExtName,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.DelTF=0 and FS_News.RecTF=1 and FS_News.AuditTF=1 and FS_News.ClassID in(" & EndClassIDList & ") order by FS_News.ID Desc"
	else
		RecNewsSql = "Select Top " & NewNumberStr & " *,FS_NewsClass.FileExtName as ClassFileExtName,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.DelTF=0 and FS_News.RecTF=1 and FS_News.AuditTF=1 order by FS_News.ID Desc"
	end if
	Set RsRecNewsObj = Conn.Execute(RecNewsSql)
	RecNews = "<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""100%"">" & Chr(13) & Chr(10)
	do while Not RsRecNewsObj.Eof
		RecNews = RecNews & "<tr>" & Chr(13) & Chr(10)
		for i = 1 to RowNumberStr
			RecNews = RecNews & "<td " & RowHeightStr & ">" & NaviPicStr & "<a " & GetCSSStyleStr(CSSStyleStr) & OpenTypeStr & "  href=""" & GetOneNewsLinkURL(RsRecNewsObj("NewsID")) & """ title="""& RsRecNewsObj("Title")&""">" & GetHTMLTitle(RsRecNewsObj("TitleStyle"),GotTopic(RsRecNewsObj("Title"),TitleNumberStr)) & "</a></td>" & Chr(13) & Chr(10)
			RsRecNewsObj.MoveNext
			if RsRecNewsObj.Eof then Exit For
		Next
		RecNews = RecNews & "</tr>" & Chr(13) & Chr(10) & CompatPicStr & Chr(13) & Chr(10)
	loop
	Set RsRecNewsObj = Nothing
	RecNews = RecNews & "</table>"
End Function
'滚动新闻
Function MarqueeNews(MarqueeNumberStr,TitleNumberStr,RowNumberStr,MarqueeWidthStr,MarqueeHeightStr,MarqueeSpeedStr,MarqueeTypeStr,DateRuleStr,OpenTypeStr,CSSStyleStr)
	Dim MarqueeSql,RsMarqueeObj,i,RikerDirection
	TitleNumberStr = GetTitleNumberStr(TitleNumberStr)
	If MarqueeTypeStr <> "" and isnull(MarqueeTypeStr)=false then
		RikerDirection = LCase(Cstr(MarqueeTypeStr))
	Else
		RikerDirection = "left"
	End If
	if MarqueeSpeedStr <> "" then MarqueeSpeedStr = " scrollamount=""" & MarqueeSpeedStr & """"
	if MarqueeTypeStr <> "" then MarqueeTypeStr = " direction=""" & MarqueeTypeStr & """"
	if MarqueeWidthStr <> "" then MarqueeWidthStr = " width=""" & MarqueeWidthStr & """"
	if MarqueeHeightStr <> "" then MarqueeHeightStr = " Height=""" & MarqueeHeightStr & """"
	OpenTypeStr = GetOpenTypeStr(OpenTypeStr)
	MarqueeSql = "Select Top " & MarqueeNumberStr & " *,FS_NewsClass.FileExtName as ClassFileExtName,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.MarqueeNews=1 and FS_News.DelTF=0 and FS_News.AuditTF=1 order by FS_News.ID Desc"
	Set RsMarqueeObj = Conn.Execute(MarqueeSql)
	MarqueeNews = "<MARQUEE" & MarqueeSpeedStr & MarqueeTypeStr & MarqueeWidthStr & MarqueeHeightStr & " onmouseover=""this.stop();"" onmouseout=""this.start();"">"
	If Cstr(RikerDirection)<>"left" and Cstr(RikerDirection)<>"right" then MarqueeNews = MarqueeNews & "<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""100%"">" & Chr(13) & Chr(10)
	do while Not RsMarqueeObj.Eof
		If Cstr(RikerDirection)<>"left" and Cstr(RikerDirection)<>"right" then MarqueeNews = MarqueeNews & "<tr>" & Chr(13) & Chr(10)
		for i = 1 to RowNumberStr
			if DateRuleStr = "" then
				If Cstr(RikerDirection)<>"left" and Cstr(RikerDirection)<>"right" then
					MarqueeNews = MarqueeNews & "<td><a class=""" & CSSStyleStr & """ href=""" & GetOneNewsLinkURL(RsMarqueeObj("NewsID")) & """ title="""& RsMarqueeObj("Title")&"""" & OpenTypeStr & ">" & GetHTMLTitle(RsMarqueeObj("TitleStyle"),GotTopic(RsMarqueeObj("Title"),TitleNumberStr)) & "</a></td>" & Chr(13) & Chr(10)
				Else
					MarqueeNews = MarqueeNews & "<a class=""" & CSSStyleStr & """ href=""" & GetOneNewsLinkURL(RsMarqueeObj("NewsID")) & """ title=""" & RsMarqueeObj("Title") &"""" & OpenTypeStr & ">" & GetHTMLTitle(RsMarqueeObj("TitleStyle"),GotTopic(RsMarqueeObj("Title"),TitleNumberStr)) & "</a>&nbsp;&nbsp;&nbsp;"
				End If
			else
				If Cstr(RikerDirection)<>"left" and Cstr(RikerDirection)<>"right" then
					MarqueeNews = MarqueeNews & "<td><a class=""" & CSSStyleStr & """" & OpenTypeStr & " href=""" & GetOneNewsLinkURL(RsMarqueeObj("NewsID")) & """ title="""& RsMarqueeObj("Title")&""""&OpenTypeStr&">" & GetHTMLTitle(RsMarqueeObj("TitleStyle"),GotTopic(RsMarqueeObj("Title"),TitleNumberStr)) & "</a>&nbsp;&nbsp;" & "<span class=""" & CSSStyleStr & """>" & DateFormat(RsMarqueeObj("AddDate"),DateRuleStr) & "</span>" & "</td>" & Chr(13) & Chr(10)
				Else
					MarqueeNews = MarqueeNews & "<a class=""" & CSSStyleStr & """" & OpenTypeStr & " href=""" & GetOneNewsLinkURL(RsMarqueeObj("NewsID")) & """ title="""& RsMarqueeObj("Title")&""""&OpenTypeStr&">" & GetHTMLTitle(RsMarqueeObj("TitleStyle"),GotTopic(RsMarqueeObj("Title"),TitleNumberStr)) & "</a>&nbsp;&nbsp;" & "<span class=""" & CSSStyleStr & """>" & DateFormat(RsMarqueeObj("AddDate"),DateRuleStr) & "</span>" & "&nbsp;&nbsp;&nbsp;"
				End If
			end if
			RsMarqueeObj.MoveNext
			if RsMarqueeObj.Eof then Exit For
		Next
		If Cstr(RikerDirection)<>"left" and Cstr(RikerDirection)<>"right" then MarqueeNews = MarqueeNews & "</tr>" & Chr(13) & Chr(10)
	loop
	Set RsMarqueeObj = Nothing
	If Cstr(RikerDirection)<>"left" and Cstr(RikerDirection)<>"right" then MarqueeNews = MarqueeNews & "</table>" & Chr(13) & Chr(10)
	MarqueeNews = MarqueeNews & "</MARQUEE>" & Chr(13) & Chr(10)
End Function
'栏目新闻列表
Function ClassNewsList(ClassListStr,NewsNumberStr,RowNumberStr,NaviPicStr,BGPicStr,RowHeightStr,CssFileStr,OpenModeStr,DetachPageStr,TitleNumberStr,DateRuleStr,DateRightStr,DateCSSStyleStr,TxtNaviStr)
	Dim ClassSql,RsClassObj,NewsSql,RsNewsObj,i
	Dim PageNum,PageIndex,LoopVar,TempClassNewsList,ClassNewsPageStr,j
	Dim TempRowNumberStr,TempDateShowStr
	TitleNumberStr = GetTitleNumberStr(TitleNumberStr)
	ClassListStr = ""
	if ClassListStr <> "" then
		ClassSql = "Select * from FS_NewsClass where ClassEName='" & ClassListStr & "'"
	else
		if RefreshType = "Class" then
			ClassSql = "Select * from FS_NewsClass where ClassID='" & RefreshID & "'"
		else
			ClassSql = ""
		end if
	end if
	if ClassSql <> "" then
		OpenModeStr = GetOpenTypeStr(OpenModeStr)
		TempRowNumberStr = GetRowSpanNumber(DateRuleStr,DateRightStr,RowNumberStr)
		BGPicStr = GetCompatPicStr(BGPicStr,DateRightStr,DateRuleStr,RowNumberStr)
		NaviPicStr = GetNewsNavitionStr(TxtNaviStr,NaviPicStr)
		if RowHeightStr <> "" then RowHeightStr = " Height=""" & RowHeightStr & """"
		Set RsClassObj = Conn.Execute(ClassSql)
		if Not RsClassObj.Eof then
			NewsNumberStr = GetTitleNumberStr(NewsNumberStr)
			Dim ClassLinkURL,ClassLinkURLName,ClassSaveExtName
			ClassLinkURL = GetOneClassLinkURL(RsClassObj("ClassEName"),RsClassObj("SaveFilePath"),RsClassObj("FileExtName"))
			ClassLinkURLName = Left(ClassLinkURL,InStrRev(ClassLinkURL,".")-1)
			ClassSaveExtName = RsClassObj("FileExtName")
			NewsSql = "Select * from FS_News Where ClassID='" & RsClassObj("ClassID") & "' and DelTF=0 and AuditTF=1 order by ID desc"
			Set RsNewsObj = Server.CreateObject(G_FS_RS)
			RsNewsObj.Open NewsSql,Conn,1,1
			if Not RsNewsObj.Eof then
				RsNewsObj.PageSize = NewsNumberStr
				PageNum = RsNewsObj.PageCount
			else
				PageNum = 0
			end if
			if (DetachPageStr = "1") and (PageNum > 1) then
				for PageIndex = 1 to PageNum
					ClassNewsPageStr = "<tr><td " & TempRowNumberStr & "><table border=""0"" width=""100%""><tr><td width=""50%"" align=""right""><br>第<font color=red>" & PageIndex&"</font>页,共<font color=red>"&PageNum & "</font>页 "
					if PageIndex = 1 then
						ClassNewsPageStr = ClassNewsPageStr & "<font face=webdings>9</font> "
						ClassNewsPageStr = ClassNewsPageStr & "<font face=webdings>7</font> "
					elseif PageIndex=2 then
						ClassNewsPageStr = ClassNewsPageStr & "<a href=""" & ClassLinkURL & """ title=首页><font face=webdings>9</font></a> "
						ClassNewsPageStr = ClassNewsPageStr & "<a href=""" & ClassLinkURL & """ title=上一页><font face=webdings>7</font></a> "
					else
						ClassNewsPageStr = ClassNewsPageStr & "<a href=""" & ClassLinkURL & """ title=首页><font face=webdings>9</font></a> "
						ClassNewsPageStr = ClassNewsPageStr & "<a href=""" & ClassLinkURLName & "_" & PageIndex-1 & "." & ClassSaveExtName & """　title=上一页><font face=webdings>7</font></a> "
					end if
					dim G
					G=0
					for j = PageIndex to PageNum
						if j = 1 then
							if j=PageIndex then
								ClassNewsPageStr = ClassNewsPageStr & "<a href=""" & ClassLinkURL & """><font color=red>[" & j & "]</font></a> "
							else
								ClassNewsPageStr = ClassNewsPageStr & "<a href=""" & ClassLinkURL & """>[" & j & "]</a> "
							end if
						else
							if j=PageIndex then
								ClassNewsPageStr = ClassNewsPageStr & "<a href=""" & ClassLinkURLName & "_" & j & "." & ClassSaveExtName & """><font color=red>[" & j & "]</font></a> "
							else
								ClassNewsPageStr = ClassNewsPageStr & "<a href=""" & ClassLinkURLName & "_" & j & "." & ClassSaveExtName & """>[" & j & "]</a> "
							end if
						end if
						G=G+1
						if G mod 10 = 0 then
					     exit for
						End if
					Next
					if PageIndex=PageNum then
						ClassNewsPageStr = ClassNewsPageStr & "<font face=webdings>8</font> "
					else
						ClassNewsPageStr = ClassNewsPageStr & "<a href=""" & ClassLinkURLName & "_" & PageIndex+1 & "." & ClassSaveExtName & """  title=下一页><font face=webdings>8</font></a> "
					end if
					if PageIndex=PageNum then
						ClassNewsPageStr = ClassNewsPageStr & "<font face=webdings>:</font> "
					else
						ClassNewsPageStr = ClassNewsPageStr & "<a href=""" & ClassLinkURLName & "_"&PageNum&"." & ClassSaveExtName & """ title=最后一页><font face=webdings>:</font></a> "
					end if
					ClassNewsPageStr = ClassNewsPageStr & "</td></tr></table></td></tr>"
					RsNewsObj.AbsolutePage = PageIndex
					TempClassNewsList = "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">" & Chr(13) & Chr(10)
					Dim TempAlreadyShow
					TempAlreadyShow = 1
					for LoopVar = 1 to RsNewsObj.PageSize
						if TempAlreadyShow > RsNewsObj.PageSize then Exit For
						if RsNewsObj.Eof then Exit For
						TempClassNewsList = TempClassNewsList & "<tr>" & Chr(13) & Chr(10)
						for i = 1 to RowNumberStr
							TempAlreadyShow = TempAlreadyShow + 1
							if DateRuleStr <> "" then
								if DateRightStr = "Left" then
									TempDateShowStr = "&nbsp;&nbsp;<span " & GetCSSStyleStr(DateCSSStyleStr) & ">" & DateFormat(RsNewsObj("AddDate"),DateRuleStr) & "</span>"
								elseif DateRightStr = "Center" then
									TempDateShowStr = "<td align=""center""><span " & GetCSSStyleStr(DateCSSStyleStr) & ">" & DateFormat(RsNewsObj("AddDate"),DateRuleStr) & "</span>" & "</td>"& Chr(13) & Chr(10)
								elseif DateRightStr = "Right" then
									TempDateShowStr = "<td align=""Right""><span " & GetCSSStyleStr(DateCSSStyleStr) & ">" & DateFormat(RsNewsObj("AddDate"),DateRuleStr) & "</span>" & "</td>"& Chr(13) & Chr(10)
								else
									TempDateShowStr = "&nbsp;&nbsp;<span " & GetCSSStyleStr(DateCSSStyleStr) & ">" & DateFormat(RsNewsObj("AddDate"),DateRuleStr) & "</span>"
								end if
							else
								TempDateShowStr = ""
							end if
							if DateRightStr = "Center" OR  DateRightStr = "Right" then 
								TempClassNewsList = TempClassNewsList & "<td " & RowHeightStr & ">" & NaviPicStr & "<a " & GetCSSStyleStr(CssFileStr) & OpenModeStr & "  href=""" & GetOneNewsLinkURL(RsNewsObj("NewsID")) & """ title="""& RsNewsObj("Title")&""">" & GetHTMLTitle(RsNewsObj("TitleStyle"),GotTopic(RsNewsObj("Title"),TitleNumberStr)) & "</a></td>" & TempDateShowStr & Chr(13) & Chr(10)
							else
								TempClassNewsList = TempClassNewsList & "<td " & RowHeightStr & ">" & NaviPicStr & "<a " & GetCSSStyleStr(CssFileStr) & OpenModeStr & "  href=""" & GetOneNewsLinkURL(RsNewsObj("NewsID")) & """ title="""& RsNewsObj("Title")&""">" & GetHTMLTitle(RsNewsObj("TitleStyle"),GotTopic(RsNewsObj("Title"),TitleNumberStr)) & "</a>" & TempDateShowStr & "</td>" & Chr(13) & Chr(10)
							end if
							RsNewsObj.MoveNext
							if TempAlreadyShow > RsNewsObj.PageSize then Exit For
							if RsNewsObj.Eof then Exit For
						Next
						TempClassNewsList = TempClassNewsList & "</tr>" & Chr(13) & Chr(10)
						TempClassNewsList = TempClassNewsList & BGPicStr & Chr(13) & Chr(10)
					Next
					TempClassNewsList = TempClassNewsList & ClassNewsPageStr & Chr(13) & Chr(10)
					TempClassNewsList = TempClassNewsList & "</table>"
					if ClassNewsList = "" then
						ClassNewsList = TempClassNewsList
					else
						ClassNewsList = ClassNewsList & "$$$" & TempClassNewsList
					end if
				Next
			else
				ClassNewsList = "<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""100%"">" & Chr(13) & Chr(10)
				do while Not RsNewsObj.Eof
					ClassNewsList = ClassNewsList & "<tr>" & Chr(13) & Chr(10)
					for i = 1 to RowNumberStr
						if DateRuleStr <> "" then
							if DateRightStr = "Left" then
								TempDateShowStr = "&nbsp;&nbsp;<span " & GetCSSStyleStr(DateCSSStyleStr) & ">" & DateFormat(RsNewsObj("AddDate"),DateRuleStr) & "</span>"
							elseif DateRightStr = "Center" then
								TempDateShowStr = "<td align=""center""><span" & GetCSSStyleStr(DateCSSStyleStr) & ">" & DateFormat(RsNewsObj("AddDate"),DateRuleStr) & "</span>" & "</td>"& Chr(13) & Chr(10)
							elseif DateRightStr = "Right" then
								TempDateShowStr = "<td align=""Right""><span " & GetCSSStyleStr(DateCSSStyleStr) & ">" & DateFormat(RsNewsObj("AddDate"),DateRuleStr) & "</span>" & "</td>"& Chr(13) & Chr(10)
							else
								TempDateShowStr = "&nbsp;&nbsp;<span " & GetCSSStyleStr(DateCSSStyleStr) & ">" & DateFormat(RsNewsObj("AddDate"),DateRuleStr) & "</span>"
							end if
						else
							TempDateShowStr = ""
						end if
						if DateRightStr = "Center" OR  DateRightStr = "Right" then 
							ClassNewsList = ClassNewsList & "<td " & RowHeightStr & ">" & NaviPicStr & "<a " & GetCSSStyleStr(CssFileStr) & OpenModeStr & "  href=""" & GetOneNewsLinkURL(RsNewsObj("NewsID")) & """>" & GetHTMLTitle(RsNewsObj("TitleStyle"),GotTopic(RsNewsObj("Title"),TitleNumberStr)) & "</a></td>" & TempDateShowStr & Chr(13) & Chr(10)
						else
							ClassNewsList = ClassNewsList & "<td " & RowHeightStr & ">" & NaviPicStr & "<a " & GetCSSStyleStr(CssFileStr) & OpenModeStr & "  href=""" & GetOneNewsLinkURL(RsNewsObj("NewsID")) & """>" & GetHTMLTitle(RsNewsObj("TitleStyle"),GotTopic(RsNewsObj("Title"),TitleNumberStr)) & "</a>" & TempDateShowStr & "</td>" & Chr(13) & Chr(10)
						end if
						RsNewsObj.MoveNext
						if RsNewsObj.Eof then Exit For
					Next
					ClassNewsList = ClassNewsList & "</tr>" & Chr(13) & Chr(10)
					ClassNewsList = ClassNewsList & BGPicStr & Chr(13) & Chr(10)
				Loop
				ClassNewsList = ClassNewsList & "</table>"
			end if
			ClassNewsList = Split(ClassNewsList,"$$$")
			Set RsNewsObj = Nothing
		else
			ClassNewsList = Array("")
		end if
		Set RsClassObj = Nothing
	else
		ClassNewsList = Array("")
	end if
End Function
'一般搜索
Function Search()
	Search = "<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""100%"">" & Chr(13) & Chr(10)
	Search = Search & "<form method=""post"" action=""" & AvailableDoMain & "/" & "search.asp"">" & Chr(13) & Chr(10)
	Search = Search & "<tr><td>" & Chr(13) & Chr(10)
	Search = Search & "<input type=""text"" name=""keyword"" size=""10"">&nbsp;" & Chr(13) & Chr(10)
	Search = Search & "<select name=""Condition""><option value=""title"">名称</option>&nbsp;&nbsp;<option value=""content"">全文</option><option value=""author"">作者/公司</option></select>" & Chr(13) & Chr(10)
	Search = Search & "<select name=""Types""><option value=""News"">信息</option>&nbsp;&nbsp;<option value=""DownLoad"">下载</option></select>" & Chr(13) & Chr(10)
	Search = Search & "&nbsp;&nbsp;<input type=""submit"" name=""Submit"" value="" 搜 索 "">" & Chr(13) & Chr(10)
	Search = Search & "</td></tr>" & Chr(13) & Chr(10)
	Search = Search & "</form>" & Chr(13) & Chr(10)
	Search = Search & "</table>" & Chr(13) & Chr(10)
End Function
'高级搜索
Function AdvancedSearch()
	Dim StrAdminDir
	If SysRootDir="" then
		StrAdminDir="/"&left(AdminDir,instr(1,AdminDir,"/"))
	Else
		StrAdminDir="/"&SysRootDir&"/"&left(AdminDir,instr(1,AdminDir,"/"))
	End If
	AdvancedSearch = "<script language=""JavaScript"" src="""&StrAdminDir&"SysJS/PublicJS.js""></script>"
	AdvancedSearch = AdvancedSearch &"<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""100%"">" & Chr(13) & Chr(10)
	AdvancedSearch = AdvancedSearch & "<form name=""AdvanceForm"" method=""post"" action=""" & AvailableDoMain & "/AdvanceSearch.asp"">" & Chr(13) & Chr(10)
	AdvancedSearch = AdvancedSearch & "<tr><td>" & Chr(13) & Chr(10)
	AdvancedSearch = AdvancedSearch & "<input type=""text"" name=""keyword"" size=""10"">&nbsp;" & Chr(13) & Chr(10)
	AdvancedSearch = AdvancedSearch & "<select name=""SClass"" id=""SClass""><option value="""" selected>选择类别</option>" & ClassList & "</select>" & Chr(13) & Chr(10)
	AdvancedSearch = AdvancedSearch & "开始时间&nbsp;<input type=""text"" size=""10"" name=""BeginDate"" readonly>"
	AdvancedSearch = AdvancedSearch & "<input type=""button"" name=""Submit4"" value=""选择"" onClick=""OpenWindowAndSetValue('"&StrAdminDir&"FunPages/SelectDate.asp',280,110,window,document.AdvanceForm.BeginDate);document.AdvanceForm.BeginDate.focus();""> "
	AdvancedSearch = AdvancedSearch & "&nbsp;结束时间&nbsp;<input type=""text"" size=""10"" name=""EndDate"" readonly>" & Chr(13) & Chr(10)
	AdvancedSearch = AdvancedSearch & "<input type=""button"" name=""Submit4"" value=""选择"" onClick=""OpenWindowAndSetValue('"&StrAdminDir&"FunPages/SelectDate.asp',280,110,window,document.AdvanceForm.EndDate);document.AdvanceForm.EndDate.focus();""> "
	AdvancedSearch = AdvancedSearch & "<select name=""Condition""><option value=""title"">名称</option><option value=""content"">全文</option><option value=""author"">作者/公司</option></select>" & Chr(13) & Chr(10)
	AdvancedSearch = AdvancedSearch & "<select name=""Types""><option value=""News"">信息</option><option value=""DownLoad"">下载</option></select>" & Chr(13) & Chr(10)
	AdvancedSearch = AdvancedSearch & "<input type=""submit"" name=""Submit"" value="" 搜 索 "">" & Chr(13) & Chr(10)
	AdvancedSearch = AdvancedSearch & "</td></tr>" & Chr(13) & Chr(10)
	AdvancedSearch = AdvancedSearch & "</form>" & Chr(13) & Chr(10)
	AdvancedSearch = AdvancedSearch & "</table>" & Chr(13) & Chr(10)
End Function
'用户登陆
Function UseLogin()
	UseLogin = "<iframe src=""" & Dummy_Riker & "/Users/UserIndex.asp"" frameborder=""0"" width=""100%"" scrollong=""no""></iframe>"
End Function
'信息统计
Function InfoStat(ClassListStr,ShowModeStr,CssFileStr)
	Dim ClassSql,RsClassObj,NewsSql,RsNewsObj,TempClassID,TempClassIDArray,AllClassID
	Dim UserSql,RsUserObj
	Dim ClassNum,NewsNum,UserNum
	if ClassListStr = "" then
		if RefreshType = "Class" then
			TempClassID = ChildClassIDList(RefreshID)
			AllClassID = "'" & RefreshID & "'" & TempClassID
			TempClassIDArray = Split(AllClassID,",")
			ClassNum = UBound(TempClassIDArray)
			NewsSql = "Select Count(ID) from FS_News where ClassID in (" & AllClassID & ")"
			Set RsNewsObj = Conn.Execute(NewsSql)
			NewsNum = RsNewsObj(0)
			Set RsNewsObj = Nothing
		elseif RefreshType = "Index" then
			ClassSql = "Select Count(ID) from FS_NewsClass"
			Set RsClassObj = Conn.Execute(ClassSql)
			ClassNum = RsClassObj(0)
			Set RsClassObj = Nothing
			NewsSql = "Select Count(ID) from FS_News"
			Set RsNewsObj = Conn.Execute(NewsSql)
			NewsNum = RsNewsObj(0)
			Set RsNewsObj = Nothing
		elseif RefreshType = "News" then
			ClassSql = "Select ClassID from FS_News where ClassID='" & RefreshID & "' order by ID desc"
			Set RsClassObj = Conn.Execute(ClassSql)
			if Not RsClassObj.Eof then
				TempClassID = ChildClassIDList(RsClassObj("ClassID"))
				AllClassID = "'" & RsClassObj("ClassID") & "'" & TempClassID
				TempClassIDArray = Split(AllClassID,",")
				ClassNum = UBound(TempClassIDArray)
				NewsSql = "Select Count(ID) from FS_News where ClassID in (" & AllClassID & ")"
				Set RsNewsObj = Conn.Execute(NewsSql)
				NewsNum = RsNewsObj(0)
				Set RsNewsObj = Nothing
			else
				ClassNum = 0
				NewsNum = 0
			end if
			Set RsClassObj = Nothing
		elseif RefreshType = "DownLoad" then
			ClassSql = "Select ClassID from FS_DownLoad where ClassID='" & RefreshID & "' order by ID desc"
			Set RsClassObj = Conn.Execute(ClassSql)
			if Not RsClassObj.Eof then
				TempClassID = ChildClassIDList(RsClassObj("ClassID"))
				AllClassID = "'" & RsClassObj("ClassID") & "'" & TempClassID
				TempClassIDArray = Split(AllClassID,",")
				ClassNum = UBound(TempClassIDArray)
				NewsSql = "Select Count(ID) from FS_News where ClassID in (" & AllClassID & ")"
				Set RsNewsObj = Conn.Execute(NewsSql)
				NewsNum = RsNewsObj(0)
				Set RsNewsObj = Nothing
			else
				ClassNum = 0
				NewsNum = 0
			end if
			Set RsClassObj = Nothing
		else
			ClassSql = ""
		end if
	else
		ClassSql = "Select ClassID from FS_NewsClass where ClassEName='" & ClassListStr & "' order by AddTime,Orders desc"
		Set RsClassObj = Conn.Execute(ClassSql)
		if Not RsClassObj.Eof then
			TempClassID = ChildClassIDList(RsClassObj("ClassID"))
			AllClassID = "'" & RsClassObj("ClassID") & "'" & TempClassID
			TempClassIDArray = Split(AllClassID,",")
			ClassNum = UBound(TempClassIDArray)
			NewsSql = "Select Count(ID) from FS_News where ClassID in (" & AllClassID & ")"
			Set RsNewsObj = Conn.Execute(NewsSql)
			NewsNum = RsNewsObj(0)
			Set RsNewsObj = Nothing
		else
			ClassNum = 0
			NewsNum = 0
		end if
		Set RsClassObj = Nothing
	end if
	UserSql = "Select Count(ID) from FS_Members"
	Set RsUserObj = Conn.Execute(UserSql)
	UserNum = RsUserObj(0)
	Set RsUserObj = Nothing
	InfoStat = "<table Class=""" & CssFileStr & """ cellpadding=""0"" cellspacing=""0"" border=""0"" width=""100%"">" & Chr(13) & Chr(10)
	if ShowModeStr = "1" then
		InfoStat = InfoStat & "<tr>" & Chr(13) & Chr(10)
		InfoStat = InfoStat & "<td>" & Chr(13) & Chr(10)
		InfoStat = InfoStat & "栏目数量&nbsp;&nbsp;" & ClassNum  & Chr(13) & Chr(10)
		InfoStat = InfoStat & "</td>" & Chr(13) & Chr(10)
		InfoStat = InfoStat & "<td>" & Chr(13) & Chr(10)
		InfoStat = InfoStat & "新闻数量&nbsp;&nbsp;" & NewsNum  & Chr(13) & Chr(10)
		InfoStat = InfoStat & "</td>" & Chr(13) & Chr(10)
		InfoStat = InfoStat & "<td>" & Chr(13) & Chr(10)
		InfoStat = InfoStat & "会员数量&nbsp;&nbsp;" & UserNum  & Chr(13) & Chr(10)
		InfoStat = InfoStat & "</td>" & Chr(13) & Chr(10)
		InfoStat = InfoStat & "</tr>" & Chr(13) & Chr(10)
	else
		InfoStat = InfoStat & "<tr>" & Chr(13) & Chr(10)
		InfoStat = InfoStat & "<td>" & "栏目数量&nbsp;&nbsp;" & ClassNum & "</td>" & Chr(13) & Chr(10)
		InfoStat = InfoStat & "</tr>" & Chr(13) & Chr(10)
		InfoStat = InfoStat & "<tr>" & Chr(13) & Chr(10)
		InfoStat = InfoStat & "<td>" & "新闻数量&nbsp;&nbsp;" & NewsNum & "</td>" & Chr(13) & Chr(10)
		InfoStat = InfoStat & "</tr>" & Chr(13) & Chr(10)
		InfoStat = InfoStat & "<tr>" & Chr(13) & Chr(10)
		InfoStat = InfoStat & "<td>" & "会员数量&nbsp;&nbsp;" & UserNum & "</td>" & Chr(13) & Chr(10)
		InfoStat = InfoStat & "</tr>" & Chr(13) & Chr(10)
	end if
	InfoStat = InfoStat & "</table>"
End Function
'相关新闻
Function RelateNews(NewsNumberStr,TitleNumberStr,RowNumberStr,NaviPicStr,CompatPicStr,CSSStyleStr,TxtNaviStr) 
	TitleNumberStr = GetTitleNumberStr(TitleNumberStr)
	if RefreshType = "News" then 
		Dim RsRelateNewsObj,RsSearchObj,RelateNewsSql,SpecialID,RSpecialIDArray,i,OldSpecialID
		CompatPicStr = GetCompatPicStr(CompatPicStr,"","",RowNumberStr)
		NaviPicStr = GetNewsNavitionStr(TxtNaviStr,NaviPicStr)
		RelateNewsSql = "Select SpecialID,KeyWords from FS_News where NewsID='" & RefreshID & "' order by ID desc"
		Set RsSearchObj = Conn.Execute(RelateNewsSql)
		if Not RsSearchObj.Eof then
			OldSpecialID = RsSearchObj("SpecialID")
			If RsSearchObj("KeyWords") <> "" and isnull(RsSearchObj("KeyWords"))=false then
				Dim KeyWordsStr,KeyWordsArray,SqlKeyWordStr,TRiker_j
				SqlKeyWordStr = ""
				KeyWordsStr = RsSearchObj("KeyWords")
				If KeyWordsStr<>"" and isnull(KeyWordsStr)=false then
					KeyWordsArray = split(KeyWordsStr,",")
					For TRiker_j = 0 to UBound(KeyWordsArray)
						If SqlKeyWordStr = "" then
							SqlKeyWordStr = "KeyWords like '%"&KeyWordsArray(TRiker_j)&"%' "
						Else
							SqlKeyWordStr = SqlKeyWordStr & "or KeyWords like '%"&KeyWordsArray(TRiker_j)&"%' "
						End If
					Next
				Else
					SqlKeyWordStr = "1=0"
				End If
				RelateNewsSql = "Select Top " & NewsNumberStr & " *,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and (" & SqlKeyWordStr & ") and FS_News.DelTF=0 and FS_News.AuditTF=1 order by FS_News.ID Desc"
				Set RsRelateNewsObj = Conn.Execute(RelateNewsSql)
				if Not RsRelateNewsObj.Eof then
					RelateNews = "<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""100%"">" & Chr(13) & Chr(10)
					do while Not RsRelateNewsObj.Eof
						RelateNews = RelateNews & "<tr>" & Chr(13) & Chr(10)
						for i =1 to RowNumberStr
							if RsRelateNewsObj("NewsID") <> RefreshID then RelateNews = RelateNews & "<td>" & NaviPicStr & "<a " & GetCSSStyleStr(CSSStyleStr) & " href= "& GetOneNewsLinkURL(RsRelateNewsObj("NewsID")) & " title="""& RsRelateNewsObj("Title")&""">"& GetHTMLTitle(RsRelateNewsObj("TitleStyle"),GotTopic(RsRelateNewsObj("Title"),TitleNumberStr)) & "</a></td>"
							RsRelateNewsObj.MoveNext
							if RsRelateNewsObj.Eof then Exit For
						next
						RelateNews = RelateNews & "</tr>" & Chr(13) & Chr(10) & CompatPicStr & Chr(13) & Chr(10)
					Loop
					RelateNews = RelateNews & "</table>"
				else
					RelateNews = ""
				end if
				RsRelateNewsObj.Close
				Set RsRelateNewsObj = Nothing
			else
				RelateNews = ""
			end if
		else
			RelateNews = ""
		end if
		Set RsSearchObj = Nothing
	else
		RelateNews = ""
	end if
End Function

Function RelateSpecialNews(NewsNumberStr,TitleNumberStr,RowNumberStr,NaviPicStr,CompatPicStr,CSSStyleStr,TxtNaviStr)
	TitleNumberStr = GetTitleNumberStr(TitleNumberStr)
	if RefreshType = "News" then 
		Dim RsRelateNewsObj,RsSearchObj,RelateNewsSql,SpecialID,RSpecialIDArray,i,OldSpecialID
		Dim RelateSpecialName,RelateSpecialNameStr,SpecialSaveFilePath
		RelateSpecialName = ""
		CompatPicStr = GetCompatPicStr(CompatPicStr,"","",RowNumberStr)
		NaviPicStr = GetNewsNavitionStr(TxtNaviStr,NaviPicStr)
		RelateNewsSql = "Select SpecialID from FS_News where NewsID='" & RefreshID & "' order by ID desc"
		Set RsSearchObj = Conn.Execute(RelateNewsSql)
		if Not RsSearchObj.Eof then
			OldSpecialID = RsSearchObj("SpecialID")
			if (Not IsNull(OldSpecialID)) And (OldSpecialID <> "") then
				Dim RelateSpecialArray,SqlSpecialSearchStr,RelateSpecial_LoopVar
				RelateSpecialArray = Split(OldSpecialID)
				For RelateSpecial_LoopVar = 0 to UBound(RelateSpecialArray)
					if RelateSpecialArray(RelateSpecial_LoopVar) <> "" then
						If SqlSpecialSearchStr = "" then
							SqlSpecialSearchStr = "FS_News.SpecialID like '%" & RelateSpecialArray(RelateSpecial_LoopVar) & "%' "
						Else
							SqlSpecialSearchStr = SqlSpecialSearchStr & "or FS_News.SpecialID like '%" & RelateSpecialArray(RelateSpecial_LoopVar) & "%' "
						End If
					end if
				Next
				RelateNewsSql = "Select Top " & NewsNumberStr & " *,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and (" & SqlSpecialSearchStr & ") and FS_News.DelTF=0 and FS_News.AuditTF=1 order by FS_News.ID Desc"
				Set RsRelateNewsObj = Conn.Execute(RelateNewsSql)
				if Not RsRelateNewsObj.Eof then
					do while Not RsRelateNewsObj.Eof
						RelateSpecialNews = RelateSpecialNews & "<tr>" & Chr(13) & Chr(10)
						for i =1 to RowNumberStr
							if RsRelateNewsObj("NewsID") <> RefreshID then
								RelateSpecialNews = RelateSpecialNews & "<td>" & NaviPicStr & "<a " & GetCSSStyleStr(CSSStyleStr) & " href= "& GetOneNewsLinkURL(RsRelateNewsObj("NewsID")) & " title="""& RsRelateNewsObj("Title")&""">"& GetHTMLTitle(RsRelateNewsObj("TitleStyle"),GotTopic(RsRelateNewsObj("Title"),TitleNumberStr)) & "</a></td>"
								if RelateSpecialName = "" then
									RelateSpecialName = RelateSpecialName & RsRelateNewsObj("SpecialID")
								else
									RelateSpecialName = RelateSpecialName & "," & RsRelateNewsObj("SpecialID")
								end if
							end if
							RsRelateNewsObj.MoveNext
							if RsRelateNewsObj.Eof then Exit For
						next
						RelateSpecialNews = RelateSpecialNews & "</tr>" & Chr(13) & Chr(10) & CompatPicStr & Chr(13) & Chr(10)
					Loop
					RsRelateNewsObj.Close
					RelateSpecialName = Replace(RelateSpecialName,",,",",")
					RelateSpecialName = Replace(RelateSpecialName,",,,",",")
					if RelateSpecialName <> "" then
						RelateSpecialName = "'" & Replace(RelateSpecialName,",","','") & "'"
						Set RsRelateNewsObj = Conn.Execute("Select * from FS_Special where SpecialID in (" & RelateSpecialName & ")")
						do while Not RsRelateNewsObj.Eof
							if RsRelateNewsObj("SaveFilePath") = "/" then
								SpecialSaveFilePath = RsRelateNewsObj("SaveFilePath")
							else
								SpecialSaveFilePath = RsRelateNewsObj("SaveFilePath") & "/"
							end if
							RelateSpecialNameStr = RelateSpecialNameStr & "<a " & GetCSSStyleStr(CSSStyleStr) & " href=""" & AvailableDoMain & SpecialSaveFilePath & RsRelateNewsObj("EName") & "/index." & RsRelateNewsObj("FileExtName") & """><strong>[" & RsRelateNewsObj("CName") & "]</strong></a>&nbsp;&nbsp;"
							RsRelateNewsObj.MoveNext
						Loop
						RsRelateNewsObj.Close
					end if
					if RelateSpecialNews <> "" then
						if RelateSpecialNameStr <> "" then
							RelateSpecialNews = "<tr><td colspan=""" & RowNumberStr & """>" & NaviPicStr & RelateSpecialNameStr & "</td></tr>" & Chr(13) & Chr(10) & RelateSpecialNews & Chr(13) & Chr(10)
							RelateSpecialNews = RelateSpecialNews & "</tr>" & Chr(13) & Chr(10) & CompatPicStr & Chr(13) & Chr(10)
							RelateSpecialNews = "<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""100%"">" & Chr(13) & Chr(10) & RelateSpecialNews & "</table>" & Chr(13) & Chr(10)
						else
							RelateSpecialNews = "<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""100%"">" & Chr(13) & Chr(10) & RelateSpecialNews & "</table>" & Chr(13) & Chr(10)
						end if
					else
						RelateSpecialNews = ""
					end if
				else
					RelateSpecialNews = ""
				end if
				Set RsRelateNewsObj = Nothing
			else
				RelateSpecialNews = ""
			end if
		else
			RelateSpecialNews = ""
		end if
		Set RsSearchObj = Nothing
	else
		RelateSpecialNews = ""
	end if
End Function
'图片新闻
Function PicNews(ClassListStr,NewsNumberStr,ShowTitleStr,OpenModeStr,TitleNumberStr,RowNumStr,PicWidthStr,PicHeightStr,CssFileStr,RowSpaceStr)
	Dim PicSql,RsPicObj,TempSql,RsTempObj,i,TitleHTMLStr
	TitleNumberStr = GetTitleNumberStr(TitleNumberStr)
	OpenModeStr = GetOpenTypeStr(OpenModeStr)
	if ClassListStr = "" then
		if RefreshType = "Index" then
			PicSql = "Select Top " & NewsNumberStr & " *,FS_NewsClass.FileExtName as ClassFileExtName,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.DelTF=0 and FS_News.PicNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ID Desc"
		elseif RefreshType = "Class" then
			PicSql = "Select Top " & NewsNumberStr & " *,FS_NewsClass.FileExtName as ClassFileExtName,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.ClassID='" & RefreshID & "' and FS_News.DelTF=0 and FS_News.PicNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ID Desc"
		elseif RefreshType = "News" then
			TempSql = "Select ClassID from FS_News where NewsID='" & RefreshID & "' order by ID desc"
			Set RsTempObj = Conn.Execute(TempSql)
			if Not RsTempobj.Eof then
				PicSql = "Select Top " & NewsNumberStr & " *,FS_NewsClass.FileExtName as ClassFileExtName,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.ClassID='" & RsTempobj("ClassID") & "' and FS_News.DelTF=0 and FS_News.PicNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ID Desc"
			else
				PicSql = "Select Top " & NewsNumberStr & " *,FS_NewsClass.FileExtName as ClassFileExtName,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.DelTF=0 and FS_News.PicNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ID Desc"
			end if
		else
			PicSql = "Select Top " & NewsNumberStr & " *,FS_NewsClass.FileExtName as ClassFileExtName,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.DelTF=0 and FS_News.PicNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ID Desc"
		end if
	else
		PicSql = "Select Top " & NewsNumberStr & " *,FS_NewsClass.FileExtName as ClassFileExtName,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_NewsClass.ClassEName='" & ClassListStr & "' and FS_News.DelTF=0 and FS_News.PicNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ID Desc"
	end if
	Set RsPicObj = Conn.Execute(PicSql)
	if Not RsPicObj.Eof then
		PicNews = "<table cellpadding=""0"" cellspacing=""" & RowSpaceStr & """ border=""0"" width=""100%"">" & Chr(13) & Chr(10)
		do while Not RsPicObj.Eof
			PicNews = PicNews & "<tr>" & Chr(13) & Chr(10)
			TitleHTMLStr = "<tr>"
			for i =1 to RowNumStr
				if ShowTitleStr = "1" then
					TitleHTMLStr = TitleHTMLStr & "<td>" & Chr(13) & Chr(10) & "<a " & OpenModeStr & GetCSSStyleStr(CssFileStr) & " href=""" & GetOneNewsLinkURL(RsPicObj("NewsID")) & """ title="""& RsPicObj("Title")&""">" & GotTopic(RsPicObj("Title"),TitleNumberStr) & "</a></td>" & Chr(13) & Chr(10)
				else
					TitleHTMLStr = ""
				end if
				dim TempDoMain
				If Left(RsPicObj("PicPath"),4)="http" then TempDoMain="" else TempDoMain=AvailableDoMain
				PicNews = PicNews & "<td>" & Chr(13) & Chr(10) & "<a " & OpenModeStr & GetCSSStyleStr(CssFileStr) & " href=""" & GetOneNewsLinkURL(RsPicObj("NewsID")) & """ title="""& RsPicObj("Title")&""">" & "<img border=""0"" src=""" & TempDoMain & RsPicObj("PicPath") & """ width=""" & PicWidthStr & """ height=""" & PicHeightStr & """>" & "</a></td>" & Chr(13) & Chr(10)
				RsPicObj.MoveNext
				if RsPicObj.Eof then Exit For
			next
			PicNews = PicNews & "</tr>" & Chr(13) & Chr(10) & TitleHTMLStr & "</tr>" & Chr(13) & Chr(10)
		Loop
		PicNews = PicNews & "</table>"
	else
		PicNews = ""
	end if
End Function
'??ID???
Function ChirldClassID(ClassEnameStr)
	Dim RsChirldFunObj,ChirldClassObj
	ChirldClassID = ""
	Set RsChirldFunObj = Conn.Execute("Select ClassID from FS_NewsClass where ParentID=(Select ClassID from FS_NewsClass where ClassEName='"&ClassEnameStr&"') order by AddTime,id desc")
	Set ChirldClassObj = Conn.Execute("Select ClassID from FS_NewsClass where ClassEName='"&ClassEnameStr&"' order by ID desc")
	If Not ChirldClassObj.eof then
		ChirldClassID = ChirldClassObj("ClassID")
		If Not RsChirldFunObj.eof then
			ChirldClassID = RsChirldFunObj("ClassID")
		End If
		Do while Not RsChirldFunObj.eof
			ChirldClassID = ChirldClassID &"','"& RsChirldFunObj("ClassID")
		RsChirldFunObj.MoveNext
		Loop
		RsChirldFunObj.Close
		Set RsChirldFunObj = Nothing
	Else
		ChirldClassID = ""
	End If
	ChirldClassObj.Close
	Set ChirldClassObj = Nothing
	ChirldClassID = ChirldClassID
End Function
'焦点图片
Function FocusPic(ClassListStr,NewsNumberStr,HaveChildStr,ShowTitleStr,TitleNumberStr,OpenModeStr,RowNumStr,PicWidthStr,PicHeightStr,CssFileStr,RowSpaceStr)
	Dim PicSql,RsPicObj,TempSql,RsTempObj,i,TitleHTMLStr
	TitleNumberStr = GetTitleNumberStr(TitleNumberStr)
	NewsNumberStr = GetTitleNumberStr(NewsNumberStr)
	OpenModeStr = GetOpenTypeStr(OpenModeStr)
	If HaveChildStr <> "" and ClassListStr <> "" then ClassListStr = ChirldClassID(ClassListStr)
	if ClassListStr = "" then
		if RefreshType = "Index" then
			PicSql = "Select Top " & NewsNumberStr & " *,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.DelTF=0 and FS_News.FocusNewsTF=1 and FS_News.PicNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ID Desc"
		elseif RefreshType = "Class" then
			PicSql = "Select Top " & NewsNumberStr & " *,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.ClassID='" & RefreshID & "' and FS_News.DelTF=0 and FS_News.FocusNewsTF=1 and FS_News.PicNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ID Desc"
		elseif RefreshType = "News" then
			TempSql = "Select ClassID from FS_News where NewsID='" & RefreshID & "'"
			Set RsTempObj = Conn.Execute(TempSql)
			if Not RsTempobj.Eof then
				PicSql = "Select Top " & NewsNumberStr & " *,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.ClassID='" & RsTempobj("ClassID") & "' and FS_News.DelTF=0 and FS_News.FocusNewsTF=1 and FS_News.PicNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ID Desc"
			else
				PicSql = "Select Top " & NewsNumberStr & " *,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.DelTF=0 and FS_News.FocusNewsTF=1 and FS_News.PicNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ID Desc"
			end if
		else
			PicSql = "Select Top " & NewsNumberStr & " *,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.DelTF=0 and FS_News.FocusNewsTF=1 and FS_News.PicNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ID Desc"
		end if
	else
		PicSql = "Select Top " & NewsNumberStr & " *,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.ClassID in ('" & ClassListStr & "') and FS_News.DelTF=0 and FS_News.FocusNewsTF=1 and FS_News.PicNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ID Desc"
	end if
	Set RsPicObj = Conn.Execute(PicSql)
	if Not RsPicObj.Eof then
		FocusPic = "<table cellpadding=""0"" cellspacing="""&RowSpaceStr&""" border=""0"" width=""100%"">" & Chr(13) & Chr(10)
		do while Not RsPicObj.Eof
			FocusPic = FocusPic & "<tr>" & Chr(13) & Chr(10)
			TitleHTMLStr = "<tr>"
			for i =1 to RowNumStr
				if ShowTitleStr = "1" then
					TitleHTMLStr = TitleHTMLStr & "<td>" & Chr(13) & Chr(10) & "<a " & OpenModeStr & GetCSSStyleStr(CssFileStr) & " href=""" & GetOneNewsLinkURL(RsPicObj("NewsID")) & """ title="""& RsPicObj("Title")&""">" & GotTopic(RsPicObj("Title"),TitleNumberStr) & "</a></td>" & Chr(13) & Chr(10)
				else
					TitleHTMLStr = ""
				end if
				dim TempDoMain
				If Left(RsPicObj("PicPath"),4)="http" then TempDoMain="" else TempDoMain=AvailableDoMain
				FocusPic = FocusPic & "<td>" & Chr(13) & Chr(10) & "<a " & OpenModeStr & GetCSSStyleStr(CssFileStr) & " href=""" & GetOneNewsLinkURL(RsPicObj("NewsID")) & """ title="""& RsPicObj("Title")&""">" & "<img border=""0"" src=""" & TempDoMain & RsPicObj("PicPath") & """ width=""" & PicWidthStr & """ height=""" & PicHeightStr & """>" & "</a></td>" & Chr(13) & Chr(10)
				RsPicObj.MoveNext
				if RsPicObj.Eof then Exit For
			next
			FocusPic = FocusPic & "</tr>" & Chr(13) & Chr(10) & TitleHTMLStr & "</tr>" & Chr(13) & Chr(10)
		Loop
		FocusPic = FocusPic & "</table>"
	else
		FocusPic = ""
	end if
End Function
'推荐图片
Function RecPic(ClassListStr,NewsNumberStr,HaveChildStr,ShowTitleStr,TitleNumberStr,OpenModeStr,RowNumStr,PicWidthStr,PicHeightStr,CssFileStr,RowSpaceStr)
	Dim PicSql,RsPicObj,TempSql,RsTempObj,i,TitleHTMLStr
	TitleNumberStr = GetTitleNumberStr(TitleNumberStr)
	NewsNumberStr = GetTitleNumberStr(NewsNumberStr)
	OpenModeStr = GetOpenTypeStr(OpenModeStr)
	If HaveChildStr <> "" and ClassListStr <> "" then ClassListStr = ChirldClassID(ClassListStr)
	if ClassListStr = "" then
		if RefreshType = "Index" then
			PicSql = "Select Top " & NewsNumberStr & " *,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.DelTF=0 and FS_News.RecTF=1 and FS_News.PicNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ID Desc"
		elseif RefreshType = "Class" then
			PicSql = "Select Top " & NewsNumberStr & " *,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.ClassID='" & RefreshID & "' and FS_News.DelTF=0 and FS_News.RecTF=1 and FS_News.PicNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ID Desc"
		elseif RefreshType = "News" then
			TempSql = "Select ClassID from FS_News where NewsID='" & RefreshID & "' order by AddDate desc"
			Set RsTempObj = Conn.Execute(TempSql)
			if Not RsTempobj.Eof then
				PicSql = "Select Top " & NewsNumberStr & " *,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.ClassID='" & RsTempobj("ClassID") & "' and FS_News.DelTF=0 and FS_News.RecTF=1 and FS_News.PicNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ID Desc"
			else
				PicSql = "Select Top " & NewsNumberStr & " *,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.DelTF=0 and FS_News.RecTF=1 and FS_News.PicNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ID Desc"
			end if
		else
			PicSql = "Select Top " & NewsNumberStr & " *,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.DelTF=0 and FS_News.RecTF=1 and FS_News.PicNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ID Desc"
		end if
	else
		PicSql = "Select Top " & NewsNumberStr & " *,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.ClassID in ('" & ClassListStr & "') and FS_News.DelTF=0 and FS_News.RecTF=1 and FS_News.PicNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ID Desc"
	end if
	Set RsPicObj = Conn.Execute(PicSql)
	if Not RsPicObj.Eof then
		RecPic = "<table cellpadding=""0"" cellspacing="""&RowSpaceStr&""" border=""0"" width=""100%"">" & Chr(13) & Chr(10)
		do while Not RsPicObj.Eof
			RecPic = RecPic & "<tr>" & Chr(13) & Chr(10)
			TitleHTMLStr = "<tr>"
			for i =1 to RowNumStr
				if ShowTitleStr = "1" then
					TitleHTMLStr = TitleHTMLStr & "<td align=""center"">" & Chr(13) & Chr(10) & "<a " & OpenModeStr & GetCSSStyleStr(CssFileStr)  & " href=""" & GetOneNewsLinkURL(RsPicObj("NewsID")) & """ title="""& RsPicObj("Title")&""">" & GotTopic(RsPicObj("Title"),TitleNumberStr) & "</a></td>" & Chr(13) & Chr(10)
				else
					TitleHTMLStr = ""
				end if
				dim TempDoMain
				If Left(RsPicObj("PicPath"),4)="http" then TempDoMain="" else TempDoMain=AvailableDoMain
				RecPic = RecPic & "<td>" & Chr(13) & Chr(10) & "<a " & OpenModeStr & GetCSSStyleStr(CssFileStr) & " href=""" & GetOneNewsLinkURL(RsPicObj("NewsID")) & """ title="""& RsPicObj("Title")&""">" & "<img border=""0"" src=""" & TempDoMain & RsPicObj("PicPath") & """ width=""" & PicWidthStr & """ height=""" & PicHeightStr & """>" & "</a></td>" & Chr(13) & Chr(10)
				RsPicObj.MoveNext
				if RsPicObj.Eof then Exit For
			next
			RecPic = RecPic & "</tr>" & Chr(13) & Chr(10) & TitleHTMLStr & "</tr>" & Chr(13) & Chr(10)
		Loop
		RecPic = RecPic & "</table>"
	else
		RecPic = ""
	end if
	Set RsPicObj = Nothing
End Function
'精彩回顾
Function ClassicalNews(ClassListStr,NewsNumberStr,HaveChildStr,ShowTitleStr,TitleNumberStr,OpenModeStr,RowNumStr,PicWidthStr,PicHeightStr,CssFileStr,RowSpaceStr)
	Dim PicSql,RsPicObj,TempSql,RsTempObj,i,TitleHTMLStr
	TitleNumberStr = GetTitleNumberStr(TitleNumberStr)
	NewsNumberStr = GetTitleNumberStr(NewsNumberStr)
	OpenModeStr = GetOpenTypeStr(OpenModeStr)
	If HaveChildStr <> "" and ClassListStr <> "" then ClassListStr = ChirldClassID(ClassListStr)
	if ClassListStr = "" then
		if RefreshType = "Index" then
			PicSql = "Select Top " & NewsNumberStr & " *,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.DelTF=0 and FS_News.ClassicalNewsTF=1 and FS_News.PicNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ID Desc"
		elseif RefreshType = "Class" then
			PicSql = "Select Top " & NewsNumberStr & " *,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.ClassID='" & RefreshID & "' and FS_News.DelTF=0 and FS_News.ClassicalNewsTF=1 and FS_News.PicNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ID Desc"
		elseif RefreshType = "News" then
			TempSql = "Select ClassID from FS_News where NewsID='" & RefreshID & "'"
			Set RsTempObj = Conn.Execute(TempSql)
			if Not RsTempobj.Eof then
				PicSql = "Select Top " & NewsNumberStr & " *,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.ClassID='" & RsTempobj("ClassID") & "' and FS_News.DelTF=0 and FS_News.ClassicalNewsTF=1 and FS_News.PicNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ID Desc"
			else
				PicSql = "Select Top " & NewsNumberStr & " *,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.DelTF=0 and FS_News.ClassicalNewsTF=1 and FS_News.PicNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ID Desc"
			end if
		else
			PicSql = "Select Top " & NewsNumberStr & " *,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.DelTF=0 and FS_News.ClassicalNewsTF=1 and FS_News.PicNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ID Desc"
		end if
	else
		PicSql = "Select Top " & NewsNumberStr & " *,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.ClassID in ('" & ClassListStr & "') and FS_News.DelTF=0 and FS_News.ClassicalNewsTF=1 and FS_News.PicNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ID Desc"
	end if
	Set RsPicObj = Conn.Execute(PicSql)
	if Not RsPicObj.Eof then
		ClassicalNews = "<table cellpadding=""0"" cellspacing="""&RowSpaceStr&""" border=""0"" width=""100%"">" & Chr(13) & Chr(10)
		do while Not RsPicObj.Eof
			ClassicalNews = ClassicalNews & "<tr>" & Chr(13) & Chr(10)
			TitleHTMLStr = "<tr>"
			for i =1 to RowNumStr
				if ShowTitleStr = "1" then
					TitleHTMLStr = TitleHTMLStr & "<td>" & Chr(13) & Chr(10) & "<a " & OpenModeStr & GetCSSStyleStr(CssFileStr) & " href=""" & GetOneNewsLinkURL(RsPicObj("NewsID")) & """ title="""& RsPicObj("Title")&""">" & GotTopic(RsPicObj("Title"),TitleNumberStr) & "</a></td>" & Chr(13) & Chr(10)
				else
					TitleHTMLStr = ""
				end if
				dim TempDoMain
				If Left(RsPicObj("PicPath"),4)="http" then TempDoMain="" else TempDoMain=AvailableDoMain
				ClassicalNews = ClassicalNews & "<td>" & Chr(13) & Chr(10) & "<a " & OpenModeStr & GetCSSStyleStr(CssFileStr) & " href=""" & GetOneNewsLinkURL(RsPicObj("NewsID")) & """ title="""& RsPicObj("Title")&""">" & "<img border=""0"" src=""" & TempDoMain & RsPicObj("PicPath") & """ width=""" & PicWidthStr & """ height=""" & PicHeightStr & """>" & "</a></td>" & Chr(13) & Chr(10)
				RsPicObj.MoveNext
				if RsPicObj.Eof then Exit For
			next
			ClassicalNews = ClassicalNews & "</tr>" & Chr(13) & Chr(10) & TitleHTMLStr & "</tr>" & Chr(13) & Chr(10)
		Loop
		ClassicalNews = ClassicalNews & "</table>"
	else
		ClassicalNews = ""
	end if
	Set RsPicObj = Nothing
End Function
'精彩图片
Function ClassicalPic(ClassListStr,NewsNumberStr,HaveChildStr,ShowTitleStr,TitleNumberStr,OpenModeStr,RowNumStr,PicWidthStr,PicHeightStr,CssFileStr,RowSpaceStr)
	Dim PicSql,RsPicObj,TempSql,RsTempObj,i,TitleHTMLStr
	TitleNumberStr = GetTitleNumberStr(TitleNumberStr)
	NewsNumberStr = GetTitleNumberStr(NewsNumberStr)
	OpenModeStr = GetOpenTypeStr(OpenModeStr)
	If HaveChildStr <> "" and ClassListStr <> "" then ClassListStr = ChirldClassID(ClassListStr)
	if ClassListStr = "" then
		if RefreshType = "Index" then
			PicSql = "Select Top " & NewsNumberStr & " *,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.DelTF=0 and FS_News.ClassicalNewsTF=1 and FS_News.PicNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ClickNum,FS_News.ID Desc"
		elseif RefreshType = "Class" then
			PicSql = "Select Top " & NewsNumberStr & " *,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.ClassID='" & RefreshID & "' and FS_News.ClassicalNewsTF=1 and FS_News.DelTF=0 and FS_News.PicNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ClickNum,FS_News.ID Desc"
		elseif RefreshType = "News" then
			TempSql = "Select ClassID from FS_News where NewsID='" & RefreshID & "' order by AddDate desc"
			Set RsTempObj = Conn.Execute(TempSql)
			if Not RsTempobj.Eof then
				PicSql = "Select Top " & NewsNumberStr & " *,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.ClassID='" & RsTempobj("ClassID") & "' and FS_News.ClassicalNewsTF=1 and FS_News.DelTF=0 and FS_News.PicNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ClickNum,FS_News.ID Desc"
			else
				PicSql = "Select Top " & NewsNumberStr & " *,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.DelTF=0 and FS_News.PicNewsTF=1 and FS_News.ClassicalNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ClickNum,FS_News.ID Desc"
			end if
		else
			PicSql = "Select Top " & NewsNumberStr & " *,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.DelTF=0 and FS_News.PicNewsTF=1 and FS_News.ClassicalNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ClickNum,FS_News.ID Desc"
		end if
	else
		PicSql = "Select Top " & NewsNumberStr & " *,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.ClassID in ('" & ClassListStr & "') and FS_News.ClassicalNewsTF=1 and FS_News.DelTF=0 and FS_News.PicNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ClickNum,FS_News.ID Desc"
	end if
	Set RsPicObj = Conn.Execute(PicSql)
	if Not RsPicObj.Eof then
		ClassicalPic = "<table cellpadding=""0"" cellspacing="""&RowSpaceStr&""" border=""0"" width=""100%"">" & Chr(13) & Chr(10)
		do while Not RsPicObj.Eof
			ClassicalPic = ClassicalPic & "<tr>" & Chr(13) & Chr(10)
			TitleHTMLStr = "<tr>"
			for i =1 to RowNumStr
				if ShowTitleStr = "1" then
					TitleHTMLStr = TitleHTMLStr & "<td><div align=""center"">" & Chr(13) & Chr(10) & "<a " & OpenModeStr & GetCSSStyleStr(CssFileStr) & " href=""" & GetOneNewsLinkURL(RsPicObj("NewsID")) & """ title="""& RsPicObj("Title")&""">" & GotTopic(RsPicObj("Title"),TitleNumberStr) & "</a></div></td>" & Chr(13) & Chr(10)
				else
					TitleHTMLStr = ""
				end if
				dim TempDoMain
				If Left(RsPicObj("PicPath"),4)="http" then TempDoMain="" else TempDoMain=AvailableDoMain
				ClassicalPic = ClassicalPic & "<td>" & Chr(13) & Chr(10) & "<a " & OpenModeStr & GetCSSStyleStr(CssFileStr) & " href=""" & GetOneNewsLinkURL(RsPicObj("NewsID")) & """ title="""& RsPicObj("Title")&""">" & "<img border=""0"" src=""" & TempDoMain & RsPicObj("PicPath") & """ width=""" & PicWidthStr & """ height=""" & PicHeightStr & """>" & "</a></td>" & Chr(13) & Chr(10)
				RsPicObj.MoveNext
				if RsPicObj.Eof then Exit For
			next
			ClassicalPic = ClassicalPic & "</tr>" & Chr(13) & Chr(10) & TitleHTMLStr & "</tr>" & Chr(13) & Chr(10)
		Loop
		ClassicalPic = ClassicalPic & "</table>"
	else
		ClassicalPic = ""
	end if
End Function
'图片终极分类
Function LastClassPic(CutPageStr,NewsNumberStr,ShowTitleStr,TitleNumberStr,OpenModeStr,RowNumStr,PicWidthStr,PicHeightStr,CssFileStr,RowSpaceStr)
	Dim PicSql,RsPicObj,TempSql,RsTempObj,i,TitleHTMLStr
	Dim PageNum,PageIndex,LoopVar,TempClassPicList,ClassPicPageStr,j,TempDoMain
	TitleNumberStr = GetTitleNumberStr(TitleNumberStr)
	NewsNumberStr = GetTitleNumberStr(NewsNumberStr)
	OpenModeStr = GetOpenTypeStr(OpenModeStr)
	if RefreshType = "Class" then
		PicSql = "Select *,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.ClassID='" & RefreshID & "' and FS_News.DelTF=0 and FS_News.PicNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ID Desc"
	else
		PicSql = ""
	end if
	if PicSql <> "" then
		Set RsPicObj = Server.CreateObject(G_FS_RS)
		RsPicObj.Open PicSql,Conn,1,1
		if Not RsPicObj.Eof then
			Dim ClassLinkURL,ClassLinkURLName,ClassSaveExtName
			ClassLinkURL = GetOneClassLinkURL(RsPicObj("ClassEName"),RsPicObj("SaveFilePath"),RsPicObj("FileExtName"))
			ClassLinkURLName = Left(ClassLinkURL,InStrRev(ClassLinkURL,".")-1)
			ClassSaveExtName = RsPicObj("FileExtName")
			RsPicObj.PageSize = NewsNumberStr
			PageNum = RsPicObj.PageCount
			if (CutPageStr = "1") and (PageNum > 1) then
				for PageIndex = 1 to PageNum
					ClassPicPageStr = "<tr><td colspan=""" & RowNumStr & """><table border=""0"" width=""100%""><tr><td width=""50%"" align=""right"">本栏目共<font color=red>" & PageNum & "</font>页,当前第<font color=red>" & PageIndex & "</font>页&nbsp;&nbsp;"
					if PageIndex = 1 then
						ClassPicPageStr = ClassPicPageStr & "<font face=webdings>9</font> "
						ClassPicPageStr = ClassPicPageStr & "<font face=webdings>7</font> "
					elseif PageIndex=2 then
						ClassPicPageStr = ClassPicPageStr & "<a href=""" & ClassLinkURL & """ title=首页><font face=webdings>9</font></a> "
						ClassPicPageStr = ClassPicPageStr & "<a href=""" & ClassLinkURL & """ title=上一页><font face=webdings>7</font></a> "
					else
						ClassPicPageStr = ClassPicPageStr & "<a href=""" & ClassLinkURL & """ title=首页><font face=webdings>9</font></a> "
						ClassPicPageStr = ClassPicPageStr & "<a href=""" & ClassLinkURLName & "_" & PageIndex-1 & "." & ClassSaveExtName & """　title=上一页><font face=webdings>7</font></a> "
					end if
					dim G
					G=0
					for j = PageIndex to PageNum
						if j = 1 then
							ClassPicPageStr = ClassPicPageStr & "<a href=""" & ClassLinkURL & """>[" & j & "]</a> "
						else
							ClassPicPageStr = ClassPicPageStr & "<a href=""" & ClassLinkURLName & "_" & j & "." & ClassSaveExtName & """>[" & j & "]</a> "
						end if
						G=G+1
						if G mod 10 = 0 then exit for
					Next
					if PageIndex=PageNum then
						ClassPicPageStr = ClassPicPageStr & "<font face=webdings>8</font> "
					else
						ClassPicPageStr = ClassPicPageStr & "<a href=""" & ClassLinkURLName & "_" & PageIndex+1 & "." & ClassSaveExtName & """  title=下一页><font face=webdings>8</font></a> "
					end if
					if PageIndex=PageNum then
						ClassPicPageStr = ClassPicPageStr & "<font face=webdings>:</font> "
					else
						ClassPicPageStr = ClassPicPageStr & "<a href=""" & ClassLinkURLName & "_"& PageNum & "." & ClassSaveExtName & """ title=最后一页><font face=webdings>:</font></a> "
					end if
					ClassPicPageStr = ClassPicPageStr & "</td></tr></table></td></tr>"
					RsPicObj.AbsolutePage = PageIndex
					TempClassPicList = "<table cellpadding=""0"" border=""0"" width=""100%"">" & Chr(13) & Chr(10)
					Dim TempAlreadyShow
					TempAlreadyShow = 1
					for LoopVar = 1 to RsPicObj.PageSize
						if TempAlreadyShow > RsPicObj.PageSize then Exit For
						if RsPicObj.Eof then Exit For
						TempClassPicList = TempClassPicList & "<tr>" & Chr(13) & Chr(10)
						TitleHTMLStr = "<tr>"
						for i = 1 to RowNumStr
							TempAlreadyShow = TempAlreadyShow + 1
							if ShowTitleStr = "1" then
								TitleHTMLStr = TitleHTMLStr & "<td>" & Chr(13) & Chr(10) & "<a " & OpenModeStr & GetCSSStyleStr(CssFileStr) & " href=""" & GetOneNewsLinkURL(RsPicObj("NewsID")) & """ title="""& RsPicObj("Title")&""">" & GotTopic(RsPicObj("Title"),TitleNumberStr) & "</a></td>" & Chr(13) & Chr(10)
							else
								TitleHTMLStr = ""
							end if
							If Left(RsPicObj("PicPath"),4)="http" then TempDoMain="" else TempDoMain=AvailableDoMain
							TempClassPicList = TempClassPicList & "<td>" & Chr(13) & Chr(10) & "<a " & OpenModeStr & GetCSSStyleStr(CssFileStr) & " href=""" & GetOneNewsLinkURL(RsPicObj("NewsID")) & """ title="""& RsPicObj("Title")&""">" & "<img border=""0"" src=""" & TempDoMain & RsPicObj("PicPath") & """ width=""" & PicWidthStr & """ height=""" & PicHeightStr & """>" & "</a></td>" & Chr(13) & Chr(10)
							RsPicObj.MoveNext
							if RsPicObj.Eof then Exit For
							if TempAlreadyShow > RsPicObj.PageSize then Exit For
						next
						TempClassPicList = TempClassPicList & "</tr>" & Chr(13) & Chr(10) & TitleHTMLStr & "</tr>" & Chr(13) & Chr(10)
					Next
					TempClassPicList = TempClassPicList & ClassPicPageStr & Chr(13) & Chr(10) & "</table>"
					if LastClassPic = "" then
						LastClassPic = TempClassPicList
					else
						LastClassPic = LastClassPic & "$$$" & TempClassPicList
					end if
				Next
			else
				LastClassPic = "<table cellpadding=""0"" border=""0"" width=""100%"">" & Chr(13) & Chr(10)
				do while Not RsPicObj.Eof
					LastClassPic = LastClassPic & "<tr>" & Chr(13) & Chr(10)
					TitleHTMLStr = "<tr>"
					for i =1 to RowNumStr
						if ShowTitleStr = "1" then
							TitleHTMLStr = TitleHTMLStr & "<td>" & Chr(13) & Chr(10) & "<a " & OpenModeStr & GetCSSStyleStr(CssFileStr) & " href=""" & GetOneNewsLinkURL(RsPicObj("NewsID")) & """ title="""& RsPicObj("Title")&""">" & GotTopic(RsPicObj("Title"),TitleNumberStr) & "</a></td>" & Chr(13) & Chr(10)
						else
							TitleHTMLStr = ""
						end if
						If Left(RsPicObj("PicPath"),4)="http" then TempDoMain="" else TempDoMain=AvailableDoMain
						LastClassPic = LastClassPic & "<td>" & Chr(13) & Chr(10) & "<a " & OpenModeStr & GetCSSStyleStr(CssFileStr) & " href=""" & GetOneNewsLinkURL(RsPicObj("NewsID")) & """ title="""& RsPicObj("Title")&""">" & "<img border=""0"" src=""" & TempDoMain & RsPicObj("PicPath") & """ width=""" & PicWidthStr & """ height=""" & PicHeightStr & """>" & "</a></td>" & Chr(13) & Chr(10)
						RsPicObj.MoveNext
						if RsPicObj.Eof then Exit For
					next
					LastClassPic = LastClassPic & "</tr>" & Chr(13) & Chr(10) & TitleHTMLStr & "</tr>" & Chr(13) & Chr(10)
				Loop
				LastClassPic = LastClassPic & "</table>"
			end if
			LastClassPic = Split(LastClassPic,"$$$")
			Set RsPicObj = Nothing
		else
			LastClassPic = Array("")
		end if
	else
		LastClassPic = Array("")
	end if
End Function
'今日头条
Function TodayNews(ClassListStr,NewsNumberStr,HaveChildStr,TitleNumberStr,RowNumberStr,NaviPicStr,CompatPicStr,OpenTypeStr,CSSStyleStr,RowHeightStr,TxtNaviStr) 
	Dim LastNewsSql,RsLastNewsObj,i,PicSql,RsTempObj,TempSql
	TitleNumberStr = GetTitleNumberStr(TitleNumberStr)
	OpenTypeStr = GetOpenTypeStr(OpenTypeStr)
	CompatPicStr = GetCompatPicStr(CompatPicStr,"","",RowNumberStr)
	NaviPicStr = GetNewsNavitionStr(TxtNaviStr,NaviPicStr)
	if RowHeightStr <> "" then RowHeightStr = " Height=""" & RowHeightStr & """"
	NewsNumberStr = GetTitleNumberStr(NewsNumberStr)
	If HaveChildStr <> "" and ClassListStr <> "" then ClassListStr = ChirldClassID(ClassListStr)
	if ClassListStr = "" then
		if RefreshType = "Index" then
			PicSql = "Select Top " & NewsNumberStr & " *,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.DelTF=0 and FS_News.TodayNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ID Desc"
		elseif RefreshType = "Class" then
			PicSql = "Select Top " & NewsNumberStr & " *,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.ClassID='" & RefreshID & "' and FS_News.DelTF=0 and FS_News.TodayNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ID Desc"
		elseif RefreshType = "News" then
			TempSql = "Select ClassID from FS_News where NewsID='" & RefreshID & "' order by AddDate desc"
			Set RsTempObj = Conn.Execute(TempSql)
			if Not RsTempobj.Eof then
				PicSql = "Select Top " & NewsNumberStr & " *,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.ClassID='" & RsTempobj("ClassID") & "' and FS_News.DelTF=0 and FS_News.TodayNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ID Desc"
			else
				PicSql = "Select Top " & NewsNumberStr & " *,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.DelTF=0 and FS_News.TodayNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ID Desc"
			end if
		else
			PicSql = "Select Top " & NewsNumberStr & " *,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.DelTF=0 and FS_News.TodayNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ID Desc"
		end if
	else
		PicSql = "Select Top " & NewsNumberStr & " *,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.DelTF=0 and FS_News.TodayNewsTF=1 and FS_News.AuditTF=1 order by FS_News.ID Desc"
	end if
	Set RsLastNewsObj = Conn.Execute(PicSql)
	TodayNews = "<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""100%"">" & Chr(13) & Chr(10)
	do while Not RsLastNewsObj.Eof
		TodayNews = TodayNews & "<tr>" & Chr(13) & Chr(10)
		for i = 1 to RowNumberStr
			TodayNews = TodayNews & "<td " & RowHeightStr & ">" & NaviPicStr & "<a " & GetCSSStyleStr(CSSStyleStr) & OpenTypeStr & " href=""" & GetOneNewsLinkURL(RsLastNewsObj("NewsID")) & """ title="""& RsLastNewsObj("Title")&""">" & GetHTMLTitle(RsLastNewsObj("TitleStyle"),GotTopic(RsLastNewsObj("Title"),TitleNumberStr)) & "</a></td>" & Chr(13) & Chr(10)
			RsLastNewsObj.MoveNext
			if RsLastNewsObj.Eof then Exit For
		Next
		TodayNews = TodayNews & "</tr>" & Chr(13) & Chr(10) & CompatPicStr & Chr(13) & Chr(10)
	loop
	Set RsLastNewsObj = Nothing
	TodayNews = TodayNews & "</table>"
End Function
'幻灯片
Function FilterNews(ClassListStr,NewsNumberStr,TitleNumberStr,CssFileStr,PicWidthStr,PicHeightStr,OpenModeStr,ShowTitleStr,RowSpaceStr)
	Dim FilterSql,RsFilterObj,FilterStr,ImagesStr,TxtStr,TxtFirst,ClassSaveFilePath,LinkStr
	TitleNumberStr = GetTitleNumberStr(TitleNumberStr)
	if ClassListStr <> "" then
		FilterSql = "Select Top " & NewsNumberStr & " *,FS_NewsClass.FileExtName as ClassFileExtName,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass where FS_News.Classid=FS_NewsClass.Classid and FS_News.DelTF=0 and FS_News.FilterNews=1 and FS_News.AuditTF=1 and FS_NewsClass.ClassEName='" & ClassListStr & "' order by FS_News.ID Desc"
	else
		FilterSql = "Select Top " & NewsNumberStr & " *,FS_NewsClass.FileExtName as ClassFileExtName,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass where FS_News.Classid=FS_NewsClass.Classid and FS_News.DelTF=0 and FS_News.FilterNews=1 and FS_News.AuditTF=1 order by FS_News.ID Desc"
	end if
	Set RsFilterObj = Conn.Execute(FilterSql)
	TxtFirst=""
	if not RsFilterObj.Eof then
		Dim Temp_Num
		Temp_Num = 0
		Do While Not RsFilterObj.Eof
			Temp_Num = Temp_Num + 1
			RsFilterObj.MoveNext
		Loop
		RsFilterObj.MoveFirst
		If Temp_Num <=1 then
			Set RsFilterObj = Nothing
			FilterNews = "至少需要两条幻灯新闻才能正确显示幻灯效果"
			Set	RsFilterObj = Nothing
			Exit Function	
		End If
		if PicWidthStr <> "" then PicWidthStr = " Width=""" & PicWidthStr & """"
		if PicHeightStr <> "" then PicHeightStr = " Height=""" & PicHeightStr & """"
		if OpenModeStr <> "0" then OpenModeStr = " target=_blank"
		if CssFileStr <> "" then CssFileStr = " Class='" & CssFileStr & "'"
		do while Not RsFilterObj.Eof
			if RsFilterObj("SaveFilePath") = "/" then
				ClassSaveFilePath = RsFilterObj("SaveFilePath")
			else
				ClassSaveFilePath = RsFilterObj("SaveFilePath") & "/"
			end if
			if (Not IsNull(RsFilterObj("PicPath"))) And (RsFilterObj("PicPath") <> "") then
				if ImagesStr = "" then 
					If Instr(1,LCase(RsFilterObj("PicPath")),"http://") <> 0 then
						ImagesStr = RsFilterObj("PicPath")
					Else
						ImagesStr = AvailableDoMain  & RsFilterObj("PicPath")
					End If
					TxtStr = "<a " & CssFileStr & " href=" & GetOneNewsLinkURL(RsFilterObj("NewsID")) & " " & OpenModeStr & ">" & GotTopic(RsFilterObj("title"),TitleNumberStr)&"</a>"
					TxtFirst = "<a " & CssFileStr & " href=" & GetOneNewsLinkURL(RsFilterObj("NewsID")) & " " & OpenModeStr & ">" & GotTopic(RsFilterObj("title"),TitleNumberStr)&"</a>"
					LinkStr =  GetOneNewsLinkURL(RsFilterObj("NewsID"))
				else
					If Instr(1,LCase(RsFilterObj("PicPath")),"http://") <> 0 then
						ImagesStr = ImagesStr &","& RsFilterObj("PicPath")
					Else
						ImagesStr = ImagesStr &","& AvailableDoMain  & RsFilterObj("PicPath")
					End If
					TxtStr = TxtStr &",<a " & CssFileStr & " href=" & GetOneNewsLinkURL(RsFilterObj("NewsID")) & " " & OpenModeStr & ">" & GotTopic(RsFilterObj("title"),TitleNumberStr)&"</a>"
					LinkStr = LinkStr & "," & GetOneNewsLinkURL(RsFilterObj("NewsID"))
				end if
			end if
			RsFilterObj.MoveNext
		loop
		FilterStr="<SCRIPT language=""VBScript"">"& Chr(13)
		FilterStr = FilterStr & "Dim FileList,FileListArr,TxtList,TxtListArr,LinkList,LinkArr"& Chr(13)
		FilterStr = FilterStr & "FileList = """ & ImagesStr & """"& Chr(13)
		FilterStr = FilterStr & "LinkList = """ & LinkStr & """"& Chr(13)
		FilterStr = FilterStr & "TxtList = """ & TxtStr & """"& Chr(13)
		FilterStr = FilterStr & "FileListArr = Split(FileList,"","")"& Chr(13)
		FilterStr = FilterStr & "LinkArr = Split(LinkList,"","")"& Chr(13)
		FilterStr = FilterStr & "TxtListArr = Split(TxtList,"","")"& Chr(13)
		FilterStr = FilterStr & "Dim CanPlay"& Chr(13)
		FilterStr = FilterStr & "CanPlay = CInt(Split(Split(navigator.appVersion,"";"")(1),"" "")(2))>5"& Chr(13)
		FilterStr = FilterStr & "Dim FilterStr"& Chr(13)
		FilterStr = FilterStr & "FilterStr = ""RevealTrans(duration=2,transition=23)"""& Chr(13)
		FilterStr = FilterStr & "FilterStr = FilterStr + "";BlendTrans(duration=2)"""& Chr(13)
		FilterStr = FilterStr & "If CanPlay Then"& Chr(13)
		'FilterStr = FilterStr & "FilterStr = FilterStr + "";progid:DXImageTransform.Microsoft.Pixelate(,enabled=false,duration=2,maxSquare=25)"""& Chr(13)
		FilterStr = FilterStr & "FilterStr = FilterStr + "";progid:DXImageTransform.Microsoft.Fade(duration=2,overlap=0)"""& Chr(13)
		'FilterStr = FilterStr & "FilterStr = FilterStr + "";progid:DXImageTransform.Microsoft.RandomDissolve(duration=2)"""& Chr(13)
		'FilterStr = FilterStr & "FilterStr = FilterStr + "";progid:DXImageTransform.Microsoft.Pixelate(MaxSquare=15,Duration=1)"""& Chr(13)
		FilterStr = FilterStr & "FilterStr = FilterStr + "";progid:DXImageTransform.Microsoft.Wipe(duration=3,gradientsize=0.25,motion=reverse)"""& Chr(13)
		FilterStr = FilterStr & "Else"& Chr(13)
		FilterStr = FilterStr & "Msgbox ""幻灯片播放具有多种动态图片切换效果，但此功能需要您的浏览器为IE5.5或以上版本，否则您将只能看到部分的切换效果。"",64"& Chr(13)
		FilterStr = FilterStr & "End If"& Chr(13)
		FilterStr = FilterStr & "Dim FilterArr"& Chr(13)
		FilterStr = FilterStr & "FilterArr = Split(FilterStr,"";"")"& Chr(13)
		FilterStr = FilterStr & "Dim PlayImg_M"& Chr(13)
		FilterStr = FilterStr & "PlayImg_M = 5 * 1000  "& Chr(13)
		FilterStr = FilterStr & "Dim I"& Chr(13)

		FilterStr = FilterStr & "I = 1"& Chr(13)
		FilterStr = FilterStr & "Sub ChangeImg"& Chr(13)
		FilterStr = FilterStr & "Do While FileListArr(I)="""""& Chr(13)
		FilterStr = FilterStr & "I = I + 1"& Chr(13)
		FilterStr = FilterStr & "If I>UBound(FileListArr) Then I = 0"& Chr(13)
		FilterStr = FilterStr & "Loop"& Chr(13)
		FilterStr = FilterStr & "Dim J"& Chr(13)
		FilterStr = FilterStr & "If I>UBound(FileListArr) Then I = 0"& Chr(13)
		FilterStr = FilterStr & "Randomize"& Chr(13)
		FilterStr = FilterStr & "J = Int(Rnd * (UBound(FilterArr)+1))"& Chr(13)
		FilterStr = FilterStr & "Img.style.filter = FilterArr(J)"& Chr(13)
		FilterStr = FilterStr & "Img.filters(0).Apply"& Chr(13)
		FilterStr = FilterStr & "Img.Src = FileListArr(I)"& Chr(13)
		FilterStr = FilterStr & "Img.filters(0).play"& Chr(13)
		FilterStr = FilterStr & "Link.Href = LinkArr(I)"& Chr(13)
		If ShowTitleStr = "1" Then
			FilterStr = FilterStr & "Txt.filters(0).Apply"& Chr(13)
			FilterStr = FilterStr & "Txt.innerHTML = TxtListArr(I)"& Chr(13)
			FilterStr = FilterStr & "Txt.filters(0).play"& Chr(13)
		End If
		FilterStr = FilterStr & "I = I + 1"& Chr(13)
		FilterStr = FilterStr & "If I>UBound(FileListArr) Then I = 0"& Chr(13)
		FilterStr = FilterStr & "TempImg.Src = FileListArr(I)"& Chr(13)
		FilterStr = FilterStr & "TempLink.Href = LinkArr(I)"& Chr(13)
		FilterStr = FilterStr & "SetTimeout ""ChangeImg"", PlayImg_M,""VBScript"""& Chr(13)
		FilterStr = FilterStr & "End Sub"& Chr(13)
		FilterStr = FilterStr & "</SCRIPT>"& Chr(13)
		FilterStr = FilterStr & "<TABLE WIDTH=""100%"" height=""100%"" BORDER=""0"" CELLSPACING="""&RowSpaceStr&""" CELLPADDING=""0"">"
		FilterStr = FilterStr & "<TR ID=""NoScript"">"
		FilterStr = FilterStr & "<TD Align=""Center"" Style=""Color:White"">对不起，图片浏览功能需脚本支持，但您的浏览器已经设置了禁止脚本运行。请您在浏览器设置中调整有关安全选项。</TD>"
		FilterStr = FilterStr & "</TR>"
		FilterStr = FilterStr & "<TR Style=""Display:none"" ID=""CanRunScript""><TD HEIGHT=""100%"" Align=""Center"" vAlign=""Center""><a id=""Link"" """&OpenModeStr&"""><Img ID=""Img"" "  & PicWidthStr & PicHeightStr & " Border=""0"" ></a>"
		FilterStr = FilterStr & "</TD></TR><TR Style=""Display:none""><TD><a id=TempLink """&OpenModeStr&"""><Img ID=""TempImg"" Border=""0""></a></TD></TR>"
		If ShowTitleStr = "1" then
			FilterStr = FilterStr & "<TR><TD HEIGHT=""100%"" Align=""Center"" vAlign=""Center"">"
			FilterStr = FilterStr & "<div ID=""Txt"" style=""PADDING-LEFT: 5px; Z-INDEX: 1; FILTER: progid:DXImageTransform.Microsoft.Fade(duration=1,overlap=0); POSITION:"">"&TxtFirst&"</div>"
			FilterStr = FilterStr & "</TD></TR>"
		End If
		FilterStr = FilterStr & "</TABLE>"& Chr(13)
		FilterStr = FilterStr & "<Script Language=""VBScript"">"& Chr(13)
		FilterStr = FilterStr & "NoScript.Style.Display = ""none"""& Chr(13)
		FilterStr = FilterStr & "CanRunScript.Style.Display = """""& Chr(13)
		FilterStr = FilterStr & "Img.Src = FileListArr(0)"& Chr(13)
		FilterStr = FilterStr & "Link.Href = LinkArr(0)"& Chr(13)
		FilterStr = FilterStr & "SetTimeout ""ChangeImg"", PlayImg_M,""VBScript"""& Chr(13)
		FilterStr = FilterStr & "</Script>"& Chr(13)
	else
		FilterStr="没有幻灯图片"
	End if
	RsFilterObj.Close
	Set RsFilterObj = Nothing
	FilterNews = FilterStr
End Function
'站点地图
Function SiteMap(ClassListStr,ShowModeStr,OpenModeStr,CssFileStr)
	Dim SiteMapSql,RsSiteMapObj,AllClassID ,ClassSql,RsClassObj
	OpenModeStr = GetOpenTypeStr(OpenModeStr)
	if ClassListStr = "" then
		SiteMapSql = "Select * from FS_NewsClass where ParentID='0' and DelFlag=0 order by Orders Desc"
	else
		SiteMapSql = "Select * from FS_NewsClass where ClassEName='" & ClassListStr & "' and DelFlag=0 order by Orders Desc"
	end if
	Set RsSiteMapObj = Conn.Execute(SiteMapSql)
	if Not RsSiteMapObj.Eof then
		SiteMap = "<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""100%"">" & Chr(13) & Chr(10)
		do while Not RsSiteMapObj.Eof
			SiteMap = SiteMap & "<tr><td>" & Chr(13) & Chr(10)
			SiteMap = SiteMap & "【<b><a " & OpenModeStr & GetCSSStyleStr(CssFileStr) & " href=""" & GetOneClassLinkURL(RsSiteMapObj("ClassEName"),RsSiteMapObj("SaveFilePath"),RsSiteMapObj("FileExtName")) & """ >" & RsSiteMapObj("ClassCName") & "</a></b>】&nbsp;&nbsp;"
			AllClassID = ChildClassIDList(RsSiteMapObj("ClassID"))
			
			if AllClassID <> "" then
				if Left(AllClassID,1) = "," then
					AllClassID = Right(AllClassID,Len(AllClassID)-1)
				end if
				ClassSql = "Select * from FS_NewsClass where ClassID in (" & AllClassID & ") and DelFlag=0 order by Orders Desc"
				Set RsClassObj = Conn.Execute(ClassSql)
				do while Not RsClassObj.Eof
					if ShowModeStr=0 then SiteMap=SiteMap & "<br>"
					SiteMap = SiteMap & "[<a " & OpenModeStr & GetCSSStyleStr(CssFileStr) & " href=""" & GetOneClassLinkURL(RsClassObj("ClassEName"),RsClassObj("SaveFilePath"),RsClassObj("FileExtName")) & """ >" & RsClassObj("ClassCName") & "</a>]&nbsp;&nbsp;"
					RsClassObj.MoveNext
				Loop
				Set RsClassObj = Nothing
			end if
			RsSiteMapObj.MoveNext
		loop
		SiteMap = SiteMap & "</td>" & Chr(13) & Chr(10) & "</tr>" & Chr(13) & Chr(10) & "</table>" & Chr(13) & Chr(10)
	else
		SiteMap = ""
	end if
	Set RsSiteMapObj = Nothing
End Function
'版权信息
Function CopyRightStr()
	Dim ConfigSql,RsConfigObj
	ConfigSql = "Select CopyRight from FS_Config"
	Set RsConfigObj = Conn.Execute(ConfigSql)
	if Not RsConfigObj.Eof then
		CopyRightStr = RsConfigObj("CopyRight")
	else
		CopyRightStr = ""
	end if
	Set RsConfigObj = Nothing
End Function
'页面标题
Function PageTitle(PageTitleBegin)
	Dim TempSql,RsTempObj
	Select Case RefreshType
		Case "Index"
			PageTitle = "首页--" & PageTitleBegin
		Case "Class"
			TempSql = "Select ClassCName from FS_NewsClass where ClassID='" & RefreshID & "'"
			Set RsTempObj = Conn.Execute(TempSql)
			if Not RsTempObj.Eof then
				PageTitle = RsTempObj("ClassCName") & "--" & PageTitleBegin
			else
				PageTitle = PageTitleBegin
			end if
			Set RsTempObj = Nothing
		Case "News"
			TempSql = "Select Title from FS_News where NewsID='" & RefreshID & "' order by ID desc"
			Set RsTempObj = Conn.Execute(TempSql)
			if Not RsTempObj.Eof then
				PageTitle = RsTempObj("Title") & "--" & PageTitleBegin
			else
				PageTitle = PageTitleBegin
			end if
			Set RsTempObj = Nothing
		Case "Special"
			TempSql = "Select CName from FS_Special where SpecialID='" & RefreshID & "'"
			Set RsTempObj = Conn.Execute(TempSql)
			if Not RsTempObj.Eof then
				PageTitle = RsTempObj("CName") & "--" & PageTitleBegin
			else
				PageTitle = PageTitleBegin
			end if
			Set RsTempObj = Nothing
		Case Else
			PageTitle = ""
	End Select
End Function
'导读新闻
Function NaviReadNews(ClassListStr,SoonClassStr,PicWidthStr,PicHeightStr,NaviPicStr,BGPicStr,TitleNumberStr,CSSStyleStr,NewsNumberStr,RowNumberStr,OpenTypeStr,RowHeightStr,TxtNaviStr)
	Dim NaviSql,RsNaviObj,i
	Dim HaveRecordTF '是否有记录
	dim TemppID,TemppSql,EndClassIDList

	If ClassListStr<>"" then
		If SoonClassStr="1" then 
			TemppSql="select ClassID from FS_NewsClass where ClassEName='" & ClassListStr & "'"
			Set TemppID=conn.execute(TemppSql)
			EndClassIDList= "'" & TemppID(0) & "'" & AllChildClassIDStrList(TemppID(0))
		Else
			TemppSql="select ClassID from FS_NewsClass where ClassEName='" & ClassListStr & "'"
			Set TemppID=conn.execute(TemppSql)
			EndClassIDList="'" & TemppID(0) & "'"
		End if
	Else
		EndClassIDList=""
	end if
	HaveRecordTF = False
	TitleNumberStr = GetTitleNumberStr(TitleNumberStr)
	OpenTypeStr = GetOpenTypeStr(OpenTypeStr)
	BGPicStr = GetCompatPicStr(BGPicStr,"","",RowNumberStr)
	NaviPicStr = GetNewsNavitionStr(TxtNaviStr,NaviPicStr)
	if RowHeightStr <> "" then RowHeightStr = " Height=""" & RowHeightStr & """"
	'
	If EndClassIDList="" then 
		NaviSql = "Select *,FS_NewsClass.FileExtName as ClassFileExtName,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.PicNewsTF=1 and FS_News.RecTF=1 and FS_News.AuditTF=1 and FS_News.DelTF=0 order by FS_News.ID Desc"
	else
		NaviSql = "Select *,FS_NewsClass.FileExtName as ClassFileExtName,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.PicNewsTF=1 and FS_News.RecTF=1 and FS_News.AuditTF=1 and FS_News.DelTF=0 and FS_News.ClassID in (" & EndClassIDList & ") order by FS_News.ID Desc"
	End if
	Set RsNaviObj = Conn.Execute(NaviSql)
	NaviReadNews = "<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""100%"">" & Chr(13) & Chr(10) & "<tr>"
	if Not RsNaviObj.Eof then
		dim TempDoMain
		If Left(RsNaviObj("PicPath"),4)="http" then TempDoMain="" else TempDoMain=AvailableDoMain
		HaveRecordTF = True
		NaviReadNews = NaviReadNews & "<td>" & Chr(13) & Chr(10)
		NaviReadNews = NaviReadNews & "<a href=""" & GetOneNewsLinkURL(RsNaviObj("NewsID")) & """>" & Chr(13) & Chr(10)
		NaviReadNews = NaviReadNews & "<img border=""0"" src=""" & TempDoMain & RsNaviObj("PicPath") & """ width=""" & PicWidthStr & """ height=""" & PicHeightStr & """>"
		NaviReadNews = NaviReadNews & "</a>" & Chr(13) & Chr(10) & "</td>" & Chr(13) & Chr(10)
	end if
	NaviReadNews = NaviReadNews & "<td>" & Chr(13) & Chr(10)

	If EndClassIDList="" then
		NaviSql = "Select Top " & NewsNumberStr & " *,FS_NewsClass.FileExtName as ClassFileExtName,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.RecTF=1 and FS_News.PicNewsTF=1 and FS_News.DelTF=0 and FS_News.AuditTF=1 order by FS_News.ID DESC"
	Else
		NaviSql = "Select Top " & NewsNumberStr & " *,FS_NewsClass.FileExtName as ClassFileExtName,FS_News.FileExtName as NewsFileExtName from FS_News,FS_NewsClass Where FS_News.ClassID=FS_NewsClass.ClassID and FS_News.RecTF=1 and FS_News.PicNewsTF=1 and FS_News.DelTF=0 and FS_News.AuditTF=1 and FS_News.ClassID in (" & EndClassIDList & ") order by FS_News.ID DESC"
	End If

	Set RsNaviObj = Conn.Execute(NaviSql)
	if Not RsNaviObj.Eof then
		HaveRecordTF = True
		NaviReadNews = NaviReadNews & "<td>" & Chr(13) & Chr(10)
		NaviReadNews = NaviReadNews & "<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""100%"">" & Chr(13) & Chr(10)
		do while Not RsNaviObj.Eof
			NaviReadNews = NaviReadNews & "<tr>" & Chr(13) & Chr(10)
			for i =1 to RowNumberStr
				NaviReadNews = NaviReadNews & "<td " & RowHeightStr & ">" & NaviPicStr & "<a " & GetCSSStyleStr(CSSStyleStr) & " href= "& GetOneNewsLinkURL(RsNaviObj("NewsID")) & " " & OpenTypeStr & " title="""& RsNaviObj("Title")&""">"& GotTopic(RsNaviObj("Title"),TitleNumberStr) & "</a></td>"
				RsNaviObj.MoveNext
				if RsNaviObj.Eof then Exit For
			next
			NaviReadNews = NaviReadNews & "</tr>" & Chr(13) & Chr(10) & BGPicStr & Chr(13) & Chr(10)
		Loop
		NaviReadNews = NaviReadNews & "</table>" & Chr(13) & Chr(10) & "</td>" & Chr(13) & Chr(10)
	end if
	NaviReadNews = NaviReadNews & "</tr>" & Chr(13) & Chr(10) & "</table>" & Chr(13) & Chr(10)
	if HaveRecordTF = False then NaviReadNews = ""
End Function
'友情链接
Function FriendLink(LinkTypeStr,TitleNumberStr,PicWidthStr,PicHeightStr,LinkNumberStr,RowNumberStr,RowHeightStr,NewWindowTF)
	Dim FriendLinkSql,RsFriendLinkObj,i
	If TitleNumberStr <> "" then
		TitleNumberStr = Cint(TitleNumberStr)
	Else
		TitleNumberStr = 10
	End If
	IF NewWindowTF = "1" Then
		NewWindowTF = "target=""_blank"""
	Else
		NewWindowTF = ""
	End IF
	Select Case RefreshType
		Case "Index"
			FriendLinkSql = "Select Top " & LinkNumberStr & " * from FS_FriendLink where Address like '%1%' and Type=" & LinkTypeStr
		Case "Class"
			FriendLinkSql = "Select Top " & LinkNumberStr & " * from FS_FriendLink where Address like '%2%' and Type=" & LinkTypeStr
		Case "News"
			FriendLinkSql = "Select Top " & LinkNumberStr & " * from FS_FriendLink where Address like '%3%' and Type=" & LinkTypeStr
		Case "Special"
			FriendLinkSql = "Select Top " & LinkNumberStr & " * from FS_FriendLink where Address like '%4%' and Type=" & LinkTypeStr
		Case Else
			FriendLinkSql = ""
	End Select
	if FriendLinkSql <> "" then
		Dim TTTemp,Temp_Str
		Set RsFriendLinkObj = Conn.Execute(FriendLinkSql)
		if Not RsFriendLinkObj.Eof then
			if RowHeightStr <> "" then RowHeightStr = " Height=""" & RowHeightStr & """"
			FriendLink = "<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""100%"">"
			do while Not RsFriendLinkObj.Eof
				FriendLink = FriendLink & "</tr>"
				for i = 1 to RowNumberStr
					TTTemp = RsFriendLinkObj("Content")
					if LinkTypeStr = "0" then
						FriendLink = FriendLink & "<td " & RowHeightStr & " ><a href=""" & RsFriendLinkObj("Url") & """ "& NewWindowTF &">" & GotTopic(TTTemp,TitleNumberStr) & "</a>"
					else
						If Instr(1,LCase(TTTemp),"http://",1)<>0 then
							Temp_Str = TTTemp
						Else
							Temp_Str = AvailableDoMain & TTTemp
						End If
						FriendLink = FriendLink & "<td " & RowHeightStr & " ><a href=""" & RsFriendLinkObj("Url") & """ "& NewWindowTF &"><img border=""0"" width=""" & PicWidthStr & """ Height=""" & PicHeightStr & """ src=""" &Temp_Str & """></a>"
					end if
					RsFriendLinkObj.MoveNext
					if RsFriendLinkObj.Eof then Exit For
				Next
				FriendLink = FriendLink & "</tr>"
			Loop
			FriendLink = FriendLink & "</table>"
		else
			FriendLink = ""
		end if
		Set RsFriendLinkObj = Nothing
	else
		FriendLink = ""
	end if
End Function
'新闻点击次数
Function NewsClickNum()
	if RefreshType = "News" then
		NewsClickNum = "<script src=" & AvailableDoMain & "/" & "Click.asp?NewsID="& RefreshID &"></script>"
	else
		NewsClickNum = ""
	end if
End Function
'子栏目ID列表
Function ChildClassIDList(ClassID)
	Dim TempRs,TempStr
	Set TempRs = Conn.Execute("Select ClassID from FS_NewsClass where DelFlag=0 and ParentID = '" & ClassID & "' order by Orders Desc")
	do while Not TempRs.Eof
		ChildClassIDList = ChildClassIDList & ",'" & TempRs("ClassID") & "'"
		ChildClassIDList = ChildClassIDList & ChildClassIDList(TempRs("ClassID"))
		TempRs.MoveNext
	loop
	TempRs.Close
	Set TempRs = Nothing
End Function
'所有栏目的列表
Function ClassList()
	Dim Rs
	Set Rs = Conn.Execute("Select ClassID,ClassCName,ClassEName from FS_newsclass where ParentID = '0' and DelFlag=0 order by Orders Desc")
	do while Not Rs.Eof
		ClassList = ClassList & "<option value=" & Rs("ClassID") & "" & ">" & Rs("ClassCName") & "</option>"& chr(10) & chr(13)
		ClassList = ClassList & AllChildClassList(Rs("ClassID"),"")
		Rs.MoveNext	
	loop
	Rs.Close
	Set Rs = Nothing
End Function
Function AllChildClassList(ClassID,Temp)
	Dim TempRs,TempStr
	Set TempRs = Conn.Execute("Select ClassID,ClassCName,ChildNum,ClassEName from FS_NewsClass where ParentID = '" & ClassID & "' and DelFlag=0 order by AddTime desc ")
	TempStr = Temp & " - "
	do while Not TempRs.Eof
		if TempRs("ChildNum") = 0 then
			AllChildClassList = AllChildClassList & "<option value="&TempRs("ClassID")&"" & ">" & TempStr & TempRs("ClassCName") & "</option>"& chr(10) & chr(13)
		else
			AllChildClassList = AllChildClassList & "<option value="&TempRs("ClassID")&"" & ">" & TempStr & TempRs("ClassCName") & "</option>"& chr(10) & chr(13)
		end if
		AllChildClassList = AllChildClassList & AllChildClassList(TempRs("ClassID"),TempStr)
		TempRs.MoveNext
	loop
	TempRs.Close
	Set TempRs = Nothing
End Function
'版本信息
Private Function GetVisionStr()
	GetVisionStr = "<!--Powered by Foosun.cn and www.skyim.com-->" & Chr(13) & Chr(10)
End Function
'栏目下载
Function ClassDownLoad(ClassListStr,NewsListNumberStr,TitleNumberStr,CompatPicStr,NaviPicStr,DateRuleStr,DateRightStr,RowHeightStr,RowNumberStr,ShowClassCNNameStr,MoreLinkTypeStr,MoreLinkContentStr,CSSStyleStr,OpenTypeStr,DownListStyleStr,TxtNaviStr)
	Dim RsDownLoadObj,DownLoadSql,RsClassObj,ClassSql,AllClassID,i,ClassCNName
	TitleNumberStr = GetTitleNumberStr(TitleNumberStr)
	OpenTypeStr = GetOpenTypeStr(OpenTypeStr)
	NaviPicStr = GetNewsNavitionStr(TxtNaviStr,NaviPicStr)
	if RowHeightStr <> "" then RowHeightStr = " Height=""" & RowHeightStr & """"
	if DateRightStr <> "" then DateRightStr = " align=""" & DateRightStr & """"
	CompatPicStr = GetCompatPicStr(CompatPicStr,"","",RowNumberStr)
	if ClassListStr <> "" then
		ClassSql = "Select ClassCName,ClassEName,ClassID,SaveFilePath,FileExtName from FS_NewsClass where ClassEName='" & ClassListStr & "'"
		Set RsClassObj = Conn.Execute(ClassSql)
		if Not RsClassObj.Eof then
			if MoreLinkContentStr <> "" then
				if MoreLinkTypeStr = "1" then
					MoreLinkContentStr="<a " & OpenTypeStr & GetCSSStyleStr(CSSStyleStr) & " href=""" & GetOneClassLinkURL(RsClassObj("ClassEName"),RsClassObj("SaveFilePath"),RsClassObj("FileExtName")) & """ ><img border=0 src=""" & MoreLinkContentStr & """></a>"
				elseif MoreLinkTypeStr = "0" then
					MoreLinkContentStr = "<a " & OpenTypeStr & GetCSSStyleStr(CSSStyleStr) & " href=""" & GetOneClassLinkURL(RsClassObj("ClassEName"),RsClassObj("SaveFilePath"),RsClassObj("FileExtName")) & """ >" & MoreLinkContentStr & "</a>"
				else
					MoreLinkContentStr = ""
				end if
			end if
			AllClassID = "'" & RsClassObj("ClassID") & "'" & ChildClassIDList(RsClassObj("ClassID"))
			DownLoadSql = "Select top " & NewsListNumberStr & " *,FS_NewsClass.FileExtName as ClassFileExtName,FS_DownLoad.FileExtName as DownLoadFileExtName from FS_DownLoad,FS_NewsClass where FS_DownLoad.ClassID=FS_NewsClass.ClassID and FS_DownLoad.AuditTF=1 and FS_NewsClass.ClassID in (" & AllClassID & ") order by FS_DownLoad.ID Desc"
			ClassDownLoad = "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">" & Chr(13) & Chr(10)
			Set RsDownLoadObj = Conn.Execute(DownLoadSql)
			do while Not RsDownLoadObj.Eof
				ClassDownLoad = ClassDownLoad & "<tr>" & Chr(13) & Chr(10)
				for i = 1 to RowNumberStr
					if ShowClassCNNameStr = "1" then
						ClassCNName = "<a " & OpenTypeStr & GetCSSStyleStr(CSSStyleStr) & " href=""" & GetOneClassLinkURL(RsDownLoadObj("ClassEName"),RsDownLoadObj("SaveFilePath"),RsDownLoadObj("ClassFileExtName")) & """ >[" & GotTopic(RsDownLoadObj("ClassCName"),TitleNumberStr) & "]</a>&nbsp;"
					else
						ClassCNName = ""
					end if
					ClassDownLoad = ClassDownLoad & "<td " & RowHeightStr & ">" & GetOneDownLoadList(RsDownLoadObj,DownListStyleStr,ClassCNName,DateRuleStr,OpenTypeStr,NaviPicStr,CSSStyleStr) & "</td>" & Chr(13) & Chr(10)
					RsDownLoadObj.MoveNext
					if RsDownLoadObj.Eof then Exit For
				Next
				ClassDownLoad = ClassDownLoad & "</tr>" & Chr(13) & Chr(10)
				ClassDownLoad = ClassDownLoad & CompatPicStr
			Loop
			if MoreLinkContentStr <> "" then
				ClassDownLoad = ClassDownLoad & "<tr><td align=""right"" colspan=""" & RowNumberStr & """>" & MoreLinkContentStr & "</td></tr>" & Chr(13) & Chr(10)
			end if
			ClassDownLoad = ClassDownLoad & "</table>" & Chr(13) & Chr(10)
			Set RsDownLoadObj = Nothing
		else
			ClassDownLoad = ""
		end if
		Set RsClassObj = Nothing
	else
		ClassDownLoad = ""
	end if
End Function
'终极栏目下载
Function DownLoadList(ClassListStr,NewsNumberStr,RowNumberStr,NaviPicStr,BGPicStr,RowHeightStr,CssFileStr,OpenModeStr,DetachPageStr,TitleNumberStr,DownListStyleStr,TxtNaviStr)
	Dim RsDownLoadListObj,DownLoadListSql,RsClassObj,ClassSql,AllClassID,i,ClassSaveFilePath
	Dim PageNum,PageIndex,LoopVar,TempDownLoadList,DownLoadPageStr,j
	TitleNumberStr = GetTitleNumberStr(TitleNumberStr)
	OpenModeStr = GetOpenTypeStr(OpenModeStr)
	NaviPicStr = GetNewsNavitionStr(TxtNaviStr,NaviPicStr)
	if RowHeightStr <> "" then RowHeightStr = " Height=""" & RowHeightStr & """"
	BGPicStr = GetCompatPicStr(BGPicStr,"","",RowNumberStr)
	ClassListStr = ""
	if ClassListStr <> "" then
		ClassSql = "Select ClassCName,ClassEName,ClassID,SaveFilePath,FileExtName from FS_NewsClass where ClassEName='" & ClassListStr & "'"
	else
		if RefreshType = "Class" then
			ClassSql = "Select ClassCName,ClassEName,ClassID,SaveFilePath,FileExtName from FS_NewsClass where ClassID='" & RefreshID & "'"
		else
			ClassSql = ""
		end if
	end if
	if ClassSql <> "" then
		Set RsClassObj = Conn.Execute(ClassSql)
		if Not RsClassObj.Eof then
			Dim ClassLinkURL,ClassLinkURLName,ClassSaveExtName
			ClassLinkURL = GetOneClassLinkURL(RsClassObj("ClassEName"),RsClassObj("SaveFilePath"),RsClassObj("FileExtName"))
			ClassLinkURLName = Left(ClassLinkURL,InStrRev(ClassLinkURL,".")-1)
			ClassSaveExtName = RsClassObj("FileExtName")
			AllClassID = "'" & RsClassObj("ClassID") & "'" & ChildClassIDList(RsClassObj("ClassID"))
			DownLoadListSql = "Select *,FS_NewsClass.FileExtName as ClassFileExtName,FS_DownLoad.FileExtName as DownLoadFileExtName from FS_DownLoad,FS_NewsClass where FS_DownLoad.ClassID=FS_NewsClass.ClassID and FS_DownLoad.AuditTF=1 and FS_NewsClass.ClassID in (" & AllClassID & ") order by FS_DownLoad.ID Desc"
			Set RsDownLoadListObj = Server.CreateObject(G_FS_RS)
			RsDownLoadListObj.Open DownLoadListSql,Conn,1,1
			if Not RsDownLoadListObj.Eof then
				RsDownLoadListObj.PageSize = NewsNumberStr
				PageNum = RsDownLoadListObj.PageCount
			else
				PageNum = 0
			end if
			if (DetachPageStr = "1") and (PageNum > 1) then
				for PageIndex = 1 to PageNum
					DownLoadPageStr = "<tr><td><table border=""0"" width=""100%""><tr><td width=""50%"" align=""right"">本栏目共<font color=red>" & PageNum & "</font>页,当前在第<font color=red>" & PageIndex & "</font>页 "
					if PageIndex = 1 then
						DownLoadPageStr = DownLoadPageStr & "<font face=webdings>9</font> "
						DownLoadPageStr = DownLoadPageStr & "<font face=webdings>7</font> "
					elseif PageIndex=2 then
						DownLoadPageStr = DownLoadPageStr & "<a href=""" & ClassLinkURL & """ title=首页><font face=webdings>9</font></a> "
						DownLoadPageStr = DownLoadPageStr & "<a href=""" & ClassLinkURL & """ title=上一页><font face=webdings>7</font></a> "
					else
						DownLoadPageStr = DownLoadPageStr & "<a href=""" & ClassLinkURL & """ title=首页><font face=webdings>9</font></a> "
						DownLoadPageStr = DownLoadPageStr & "<a href=""" & ClassLinkURLName & "_"&PageIndex-1&"." & ClassSaveExtName & """　title=上一页><font face=webdings>7</font></a> "
					end if
					dim G
					G=0
					for j = PageIndex to PageNum
						if j = 1 then
							DownLoadPageStr = DownLoadPageStr & "<a href=""" & ClassLinkURL & """>[" & j & "]</a> "
						else
							DownLoadPageStr = DownLoadPageStr & "<a href=""" & ClassLinkURLName & "_" & j & "." & ClassSaveExtName & """>[" & j & "]</a> "
						end if
						G=G+1
						if G mod 10 = 0 then
					     exit for
						End if
					Next
					if PageIndex=PageNum then
						DownLoadPageStr = DownLoadPageStr & "<font face=webdings>8</font> "
					else
						DownLoadPageStr = DownLoadPageStr & "<a href=""" & ClassLinkURL & "_" & PageIndex+1 & "." & ClassSaveExtName & """  title=下一页><font face=webdings>8</font></a> "
					end if
					if PageIndex=PageNum then
						DownLoadPageStr = DownLoadPageStr & "<font face=webdings>:</font> "
					else
						DownLoadPageStr = DownLoadPageStr & "<a href=""" & ClassLinkURL & "_" & PageNum & "." & ClassSaveExtName & """ title=最后一页><font face=webdings>:</font></a> "
					end if
					DownLoadPageStr = DownLoadPageStr & "</td></tr></table></td></tr>"
					RsDownLoadListObj.AbsolutePage = PageIndex
					Dim TempAlreadyShow
					TempAlreadyShow = 1
					TempDownLoadList = "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">" & Chr(13) & Chr(10)
					for LoopVar = 1 to RsDownLoadListObj.PageSize
						if TempAlreadyShow > RsDownLoadListObj.PageSize then Exit For
						if RsDownLoadListObj.Eof then Exit For
						TempDownLoadList = TempDownLoadList & "<tr>" & Chr(13) & Chr(10)
						for i = 1 to RowNumberStr
							TempAlreadyShow = TempAlreadyShow + 1
							TempDownLoadList = TempDownLoadList & "<td " & RowHeightStr & ">" & GetOneDownLoadList(RsDownLoadListObj,DownListStyleStr,"","",OpenModeStr,NaviPicStr,CssFileStr) & "</td>" & Chr(13) & Chr(10)
							RsDownLoadListObj.MoveNext
							if RsDownLoadListObj.Eof then Exit For
							if TempAlreadyShow > RsDownLoadListObj.PageSize then Exit For
						Next
						TempDownLoadList = TempDownLoadList & "</tr>" & Chr(13) & Chr(10) & BGPicStr & Chr(13) & Chr(10)
					Next
					TempDownLoadList = TempDownLoadList & DownLoadPageStr & Chr(13) & Chr(10) & "</table>" & Chr(13) & Chr(10)
					if DownLoadList = "" then
						DownLoadList = TempDownLoadList
					else
						DownLoadList = DownLoadList & "$$$" & TempDownLoadList
					end if
				Next
			else
				DownLoadList = "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">" & Chr(13) & Chr(10)
				do while Not RsDownLoadListObj.Eof
					DownLoadList = DownLoadList & "<tr>" & Chr(13) & Chr(10)
					for i = 1 to RowNumberStr
						DownLoadList = DownLoadList & "<td " & RowHeightStr & ">" & GetOneDownLoadList(RsDownLoadListObj,DownListStyleStr,"","",OpenModeStr,NaviPicStr,CssFileStr) & "</td>" & Chr(13) & Chr(10)
						RsDownLoadListObj.MoveNext
						if RsDownLoadListObj.Eof then Exit For
					Next
					DownLoadList = DownLoadList & "</tr>" & Chr(13) & Chr(10) & BGPicStr & Chr(13) & Chr(10)
				Loop
				DownLoadList = DownLoadList & "</table>" & Chr(13) & Chr(10)
			end if
			Set RsDownLoadListObj = Nothing
			DownLoadList = Split(DownLoadList,"$$$")
		else
			DownLoadList = Array("")
		end if
		Set RsClassObj = Nothing
	else
		DownLoadList = Array("")
	end if
End Function
'最新下载
Function LastDownList(ClassListStr,NewNumberStr,TitleNumberStr,RowNumberStr,NaviPicStr,CompatPicStr,OpenTypeStr,CSSStyleStr,RowHeightStr,DownListStyleStr,TxtNaviStr)
	Dim LastDownListObj,LastDownListSql,i
	TitleNumberStr = GetTitleNumberStr(TitleNumberStr)
	OpenTypeStr = GetOpenTypeStr(OpenTypeStr)
	NaviPicStr = GetNewsNavitionStr(TxtNaviStr,NaviPicStr)
	if RowHeightStr <> "" then RowHeightStr = " Height=""" & RowHeightStr & """"
	CompatPicStr =GetCompatPicStr(CompatPicStr,"","",RowNumberStr)
	'调用子栏目修改2005年3月16日
	Dim TempRsHotDownListObj,AllClassID 
	Set TempRsHotDownListObj = Conn.Execute("Select ClassID from FS_NewsClass where ClassEName='" & ClassListStr & "'")
	if Not TempRsHotDownListObj.Eof then
		AllClassID = "'" & TempRsHotDownListObj("ClassID") & "'" & ChildClassIDList(TempRsHotDownListObj("ClassID"))
	else
		LastDownList = ""
		Set TempRsHotDownListObj = Nothing
'		Exit Function
	end if
	Set TempRsHotDownListObj = Nothing
	'调用子栏目修改2005年3月16日
	if ClassListStr <> "" then
		LastDownListSql = "Select top " & NewNumberStr & " *,FS_NewsClass.FileExtName as ClassFileExtName,FS_DownLoad.FileExtName as DownLoadFileExtName from FS_DownLoad,FS_NewsClass where FS_DownLoad.ClassID=FS_NewsClass.ClassID and FS_DownLoad.AuditTF=1 and FS_NewsClass.ClassID in (" & AllClassID & ") order by FS_DownLoad.ID Desc"
	else
		LastDownListSql = "Select top " & NewNumberStr & " *,FS_NewsClass.FileExtName as ClassFileExtName,FS_DownLoad.FileExtName as DownLoadFileExtName from FS_DownLoad,FS_NewsClass where FS_DownLoad.ClassID=FS_NewsClass.ClassID and FS_DownLoad.AuditTF=1 order by FS_DownLoad.ID Desc"
	end if
	Set LastDownListObj = Conn.Execute(LastDownListSql)
	if Not LastDownListObj.Eof then
		LastDownList = "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">" & Chr(13) & Chr(10)
		do while Not LastDownListObj.Eof
			LastDownList = LastDownList & "<tr>" & Chr(13) & Chr(10)
			for i = 1 to RowNumberStr
				LastDownList = LastDownList & "<td " & RowHeightStr & ">" & GetOneDownLoadList(LastDownListObj,DownListStyleStr,"","",OpenTypeStr,NaviPicStr,CSSStyleStr) & "</td>" & Chr(13) & Chr(10)
				LastDownListObj.MoveNext
				if LastDownListObj.Eof then
					Exit For
				end if
			Next
			LastDownList = LastDownList & "</tr>" & Chr(13) & Chr(10) & CompatPicStr & Chr(13) & Chr(10)
			LastDownList = LastDownList
		Loop
		LastDownList = LastDownList & "</table>" & Chr(13) & Chr(10)
		Set LastDownListObj = Nothing
	else
		LastDownList = ""
	end if
End Function
'推荐下载
Function RecDownList(ClassListStr,NewNumberStr,TitleNumberStr,RowNumberStr,NaviPicStr,CompatPicStr,OpenTypeStr,CSSStyleStr,RowHeightStr,DownListStyleStr,TxtNaviStr)
	Dim RecDownListObj,RecDownListSql,i
	TitleNumberStr = GetTitleNumberStr(TitleNumberStr)
	OpenTypeStr = GetOpenTypeStr(OpenTypeStr)
	NaviPicStr = GetNewsNavitionStr(TxtNaviStr,NaviPicStr)
	if RowHeightStr <> "" then RowHeightStr = " Height=""" & RowHeightStr & """"
	CompatPicStr = GetCompatPicStr(CompatPicStr,"","",RowNumberStr)
	'调用子栏目修改2005年3月16日
	Dim TempRsHotDownListObj,AllClassID 
	Set TempRsHotDownListObj = Conn.Execute("Select ClassID from FS_NewsClass where ClassEName='" & ClassListStr & "'")
	if Not TempRsHotDownListObj.Eof then
		AllClassID = "'" & TempRsHotDownListObj("ClassID") & "'" & ChildClassIDList(TempRsHotDownListObj("ClassID"))
	else
		RecDownList = ""
		Set TempRsHotDownListObj = Nothing
'		Exit Function
	end if
	Set TempRsHotDownListObj = Nothing
	'调用子栏目修改2005年3月16日
	if ClassListStr <> "" then
		RecDownListSql = "Select top " & NewNumberStr & " *,FS_NewsClass.FileExtName as ClassFileExtName,FS_DownLoad.FileExtName as DownLoadFileExtName from FS_DownLoad,FS_NewsClass where FS_DownLoad.ClassID=FS_NewsClass.ClassID and FS_DownLoad.AuditTF=1 and FS_RecTF=1 and FS_NewsClass.ClassID in (" & AllClassID & ") order by FS_DownLoad.ID Desc"
	else
		RecDownListSql = "Select top " & NewNumberStr & " *,FS_NewsClass.FileExtName as ClassFileExtName,FS_DownLoad.FileExtName as DownLoadFileExtName from FS_DownLoad,FS_NewsClass where FS_DownLoad.ClassID=FS_NewsClass.ClassID and FS_DownLoad.AuditTF=1 and FS_RecTF=1 order by FS_DownLoad.ID Desc"
	end if
	Set RecDownListObj = Conn.Execute(RecDownListSql)
	if Not RecDownListObj.Eof then
		RecDownList = "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">" & Chr(13) & Chr(10)
		do while Not RecDownListObj.Eof
			RecDownList = RecDownList & "<tr>" & Chr(13) & Chr(10)
			for i = 1 to RowNumberStr
				RecDownList = RecDownList & "<td " & RowHeightStr & ">" & GetOneDownLoadList(RecDownListObj,DownListStyleStr,"","",OpenTypeStr,NaviPicStr,CSSStyleStr) & "</td>" & Chr(13) & Chr(10)
				RecDownListObj.MoveNext
				if RecDownListObj.Eof then
					Exit For
				end if
			Next
			RecDownList = RecDownList & "</tr>" & Chr(13) & Chr(10) & CompatPicStr
		Loop
		RecDownList = RecDownList & "</table>" & Chr(13) & Chr(10)
		Set RecDownListObj = Nothing
	else
		RecDownList = ""
	end if
End Function
'热点下载
Function HotDownList(ClassListStr,NewNumberStr,TitleNumberStr,RowNumberStr,NaviPicStr,CompatPicStr,OpenTypeStr,CSSStyleStr,RowHeightStr,DownListStyleStr,TxtNaviStr)
	Dim HotDownListObj,HotDownListSql,i
	TitleNumberStr = GetTitleNumberStr(TitleNumberStr)
	OpenTypeStr = GetOpenTypeStr(OpenTypeStr)
	NaviPicStr = GetNewsNavitionStr(TxtNaviStr,NaviPicStr)
	if RowHeightStr <> "" then RowHeightStr = " Height=""" & RowHeightStr & """"
	CompatPicStr = GetCompatPicStr(CompatPicStr,"","",RowNumberStr)
	'调用子栏目修改2005年3月16日
	Dim TempRsHotDownListObj,AllClassID
	Set TempRsHotDownListObj = Conn.Execute("Select ClassID from FS_NewsClass where ClassEName='" & ClassListStr & "'")
	if Not TempRsHotDownListObj.Eof then
		AllClassID = "'" & TempRsHotDownListObj("ClassID") & "'" & ChildClassIDList(TempRsHotDownListObj("ClassID"))
	else
		HotDownList = ""
		Set TempRsHotDownListObj = Nothing
'		Exit Function
	end if
	Set TempRsHotDownListObj = Nothing
	'调用子栏目修改2005年3月16日
	if ClassListStr <> "" then
		HotDownListSql = "Select top " & NewNumberStr & " *,FS_NewsClass.FileExtName as ClassFileExtName,FS_DownLoad.FileExtName as DownLoadFileExtName from FS_DownLoad,FS_NewsClass where FS_DownLoad.ClassID=FS_NewsClass.ClassID and FS_DownLoad.AuditTF=1 and FS_NewsClass.ClassID in (" & AllClassID & ") order by FS_DownLoad.ClickNum Desc"
	else
		HotDownListSql = "Select top " & NewNumberStr & " *,FS_NewsClass.FileExtName as ClassFileExtName,FS_DownLoad.FileExtName as DownLoadFileExtName from FS_DownLoad,FS_NewsClass where FS_DownLoad.ClassID=FS_NewsClass.ClassID and FS_DownLoad.AuditTF=1 order by FS_DownLoad.ClickNum Desc"
	end if
	Set HotDownListObj = Conn.Execute(HotDownListSql)
	if Not HotDownListObj.Eof then
		HotDownList = "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">" & Chr(13) & Chr(10)
		do while Not HotDownListObj.Eof
			HotDownList = HotDownList & "<tr>" & Chr(13) & Chr(10)
			for i = 1 to RowNumberStr
				HotDownList = HotDownList & "<td " & RowHeightStr & ">" & GetOneDownLoadList(HotDownListObj,DownListStyleStr,"","",OpenTypeStr,NaviPicStr,CSSStyleStr) & "</td>" & Chr(13) & Chr(10)
				HotDownListObj.MoveNext
				if HotDownListObj.Eof then Exit For
			Next
			HotDownList = HotDownList & "</tr>" & Chr(13) & Chr(10) & CompatPicStr & Chr(13) & Chr(10)
		Loop
		HotDownList = HotDownList & "</table>" & Chr(13) & Chr(10)
		Set HotDownListObj = Nothing
	else
		HotDownList = ""
	end if
End Function
'得到一条下载内容
Function GetOneDownLoadList(AlreadyDownListObj,StyleID,ClassCNName,DateRuleStr,OpenTypeStr,NaviPicStr,CSSStyleStr)
	Dim DownListSql,RsDownListObj,StyleContent
	Dim TempStr
	if Not AlreadyDownListObj.Eof then
		if StyleID <> "" then
			DownListSql = "Select * from FS_DownListStyle where ID=" & StyleID & ""
			Set RsDownListObj = Conn.Execute(DownListSql)
			if Not RsDownListObj.Eof then
				StyleContent = RsDownListObj("Content")
				if Not IsNull(AlreadyDownListObj("Name")) then
					StyleContent = Replace(StyleContent,"{DownLoad_Name}",NaviPicStr & ClassCNName & "<a " & GetCSSStyleStr(CSSStyleStr) & " href=""" & GetOneDownLoadLinkURL(AlreadyDownListObj("DownLoadID")) & """ " & OpenTypeStr & ">" & AlreadyDownListObj("Name") & "</a>")
				else
					StyleContent = Replace(StyleContent,"{DownLoad_Name}","")
				end if
				if Not IsNull(AlreadyDownListObj("Version")) then
					StyleContent = Replace(StyleContent,"{DownLoad_Version}",AlreadyDownListObj("Version"))
				else
					StyleContent = Replace(StyleContent,"{DownLoad_Version}","")
				end if
				StyleContent = Replace(StyleContent,"{DownLoad_ClickNum}","<script src=" & AvailableDoMain & "/" & "DownClick.asp?DownLoadID="& AlreadyDownListObj("DownLoadID") &"></script>")
				if Not IsNull(AlreadyDownListObj("Types")) then
					Select Case AlreadyDownListObj("Types")
						Case 1 TempStr = "图片"
						Case 2 TempStr = "文件"
						Case 3 TempStr = "程序"
						Case 4 TempStr = "Flash"
						Case 5 TempStr = "音乐"
						Case 6 TempStr = "影视"
						Case 7 TempStr = "其他"
						Case Else TempStr = ""
					End Select
					StyleContent = Replace(StyleContent,"{DownLoad_Types}",TempStr)
				else
					StyleContent = Replace(StyleContent,"{DownLoad_Types}","")
				end if
				if Not IsNull(AlreadyDownListObj("Language")) then
					Select Case AlreadyDownListObj("Language")
						Case 1 TempStr = "简体中文"
						Case 2 TempStr = "繁体中文"
						Case 3 TempStr = "英文"
						Case 4 TempStr = "法文"
						Case 5 TempStr = "日文"
						Case 6 TempStr = "德文"
						Case Else TempStr = ""
					End Select
					StyleContent = Replace(StyleContent,"{DownLoad_Language}",TempStr)
				else
					StyleContent = Replace(StyleContent,"{DownLoad_Language}","")
				end if
				if Not IsNull(AlreadyDownListObj("Accredit")) then
					Select Case AlreadyDownListObj("Accredit")
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
					StyleContent = Replace(StyleContent,"{DownLoad_Accredit}",TempStr)
				else
					StyleContent = Replace(StyleContent,"{DownLoad_Accredit}","")
				end if
				if Not IsNull(AlreadyDownListObj("FileSize")) then
					StyleContent = Replace(StyleContent,"{DownLoad_FileSize}",AlreadyDownListObj("FileSize"))
				else
					StyleContent = Replace(StyleContent,"{DownLoad_FileSize}","")
				end if
				if Not IsNull(AlreadyDownListObj("Appraise")) then
					Select Case AlreadyDownListObj("Appraise")
						Case 1 TempStr = "★"
						Case 2 TempStr = "★★"
						Case 3 TempStr = "★★★"
						Case 4 TempStr = "★★★★"
						Case 5 TempStr = "★★★★★"
						Case 6 TempStr = "★★★★★★"
						Case Else TempStr = ""
					End Select
					StyleContent = Replace(StyleContent,"{DownLoad_Appraise}",TempStr)
				else
					StyleContent = Replace(StyleContent,"{DownLoad_Appraise}","")
				end if
				if Not IsNull(AlreadyDownListObj("SystemType")) then
					StyleContent = Replace(StyleContent,"{DownLoad_SystemType}",AlreadyDownListObj("SystemType"))
				else
					StyleContent = Replace(StyleContent,"{DownLoad_SystemType}","")
				end if
				if Not IsNull(AlreadyDownListObj("EMail")) then
					StyleContent = Replace(StyleContent,"{DownLoad_EMail}",AlreadyDownListObj("EMail"))
				else
					StyleContent = Replace(StyleContent,"{DownLoad_EMail}","")
				end if
				if Not IsNull(AlreadyDownListObj("ProviderUrl")) then
					StyleContent = Replace(StyleContent,"{DownLoad_ProviderUrl}",AlreadyDownListObj("ProviderUrl"))
				else
					StyleContent = Replace(StyleContent,"{DownLoad_ProviderUrl}","")
				end if
				if Not IsNull(AlreadyDownListObj("Provider")) then
					StyleContent = Replace(StyleContent,"{DownLoad_Provider}",AlreadyDownListObj("Provider"))
				else
					StyleContent = Replace(StyleContent,"{DownLoad_Provider}","")
				end if
				if Not IsNull(AlreadyDownListObj("PassWord")) then
					StyleContent = Replace(StyleContent,"{DownLoad_PassWord}",AlreadyDownListObj("PassWord"))
				else
					StyleContent = Replace(StyleContent,"{DownLoad_PassWord}","")
				end if
				if Not IsNull(AlreadyDownListObj("AddTime")) then
					StyleContent = Replace(StyleContent,"{DownLoad_AddTime}",DateFormat(AlreadyDownListObj("AddTime"),DateRuleStr))
				else
					StyleContent = Replace(StyleContent,"{DownLoad_AddTime}","")
				end if
				if Not IsNull(AlreadyDownListObj("EditTime")) then
					StyleContent = Replace(StyleContent,"{DownLoad_EditTime}",DateFormat(AlreadyDownListObj("EditTime"),DateRuleStr))
				else
					StyleContent = Replace(StyleContent,"{DownLoad_EditTime}","")
				end if
				'===========================================
				'控制下载列表中的简介字数
				TempStr = AlreadyDownListObj("Description")
				If len(TempStr)>40 then 
					TempStr = left(TempStr,40) & "......"
				Else	
				End If
				'==========================================
				if Not IsNull(TempStr) then
					StyleContent = Replace(StyleContent,"{DownLoad_Description}",TempStr)
				else
					StyleContent = Replace(StyleContent,"{DownLoad_Description}","")
				end if
				if Not IsNull(AlreadyDownListObj("Pic")) then
					StyleContent = Replace(StyleContent,"{DownLoad_Pic}","<img border=""0"" width=""60"" height=""60"" src=""" & AlreadyDownListObj("Pic") & """>")
				else
					StyleContent = Replace(StyleContent,"{DownLoad_Pic}","")
				end if
			else
				StyleContent = "没有下载列表显示的样式"
			end if
			Set RsDownListObj = Nothing
		else
			StyleContent = "没有下载列表显示的样式"
		end if
	else
		StyleContent = ""
		Exit Function
	end if
	GetOneDownLoadList = StyleContent
End Function
'??????
Function DownInfoStat(ClassListStr,ShowModeStr,CssFileStr)
	Dim ClassSql,RsClassObj,NewsSql,RsNewsObj,TempClassID,TempClassIDArray,AllClassID
	Dim UserSql,RsUserObj
	Dim ClassNum,NewsNum,UserNum
	Dim DownLoadClickNum,TodayDownLoadNum
	Dim FSO,FolderObj,DownLoadFolderSize,sRootDir
	if SysRootDir<>"" then sRootDir="/" & SysRootDir else sRootDir=""
	Set FSO = Server.CreateObject(G_FS_FSO)
	Set FolderObj = FSO.GetFolder(Server.MapPath(sRootDir &"/" & UpFiles & "/" & DownLoadDir))
	DownLoadFolderSize = CLng(FolderObj.Size/1024) & "KB"
	Set FSO = Nothing
	Set FolderObj = Nothing
	if ClassListStr = "" then
		if RefreshType = "Class" then
			TempClassID = ChildClassIDList(RefreshID)
			AllClassID = "'" & RefreshID & "'" & TempClassID
			TempClassIDArray = Split(AllClassID,",")
			ClassNum = UBound(TempClassIDArray)
			NewsSql = "Select Count(ID) from FS_DownLoad where AuditTF=1 and ClassID in (" & AllClassID & ")"
			Set RsNewsObj = Conn.Execute(NewsSql)
			NewsNum = RsNewsObj(0)
			Set RsNewsObj = Nothing
			NewsSql = "Select Sum(ClickNum) from FS_DownLoad where AuditTF=1 and ClassID in (" & AllClassID & ")"
			Set RsNewsObj = Conn.Execute(NewsSql)
			DownLoadClickNum = RsNewsObj(0)
			Set RsNewsObj = Nothing
			If IsSqlDataBase=0 then
				NewsSql = "Select Count(ID) from FS_DownLoad where AuditTF=1 and AddTime>=#" & date() &" 00:00:00# and AddTime<=#" & date() &" 23:59:59# and ClassID in (" & AllClassID & ")"
			Else
				NewsSql = "Select Count(ID) from FS_DownLoad where AuditTF=1 and AddTime>='" & date() &" 00:00:00' and AddTime<='" & date() &" 23:59:59' and ClassID in (" & AllClassID & ")"
			End If
			Set RsNewsObj = Conn.Execute(NewsSql)
			TodayDownLoadNum = RsNewsObj(0)
			Set RsNewsObj = Nothing
		elseif RefreshType = "Index" then
			ClassSql = "Select Count(ID) from FS_NewsClass where DelFlag=0"
			Set RsClassObj = Conn.Execute(ClassSql)
			ClassNum = RsClassObj(0)
			Set RsClassObj = Nothing
			NewsSql = "Select Count(ID) from FS_DownLoad where AuditTF=1" '??  Riker
			Set RsNewsObj = Conn.Execute(NewsSql)
			NewsNum = RsNewsObj(0)
			Set RsNewsObj = Nothing
			NewsSql = "Select Sum(ClickNum) from FS_DownLoad where AuditTF=1"
			Set RsNewsObj = Conn.Execute(NewsSql)
			DownLoadClickNum = RsNewsObj(0)
			Set RsNewsObj = Nothing
			If IsSqlDataBase=0 then
				NewsSql = "Select Count(ID) from FS_DownLoad where AuditTF=1 and AddTime>=#" & date() &" 00:00:00# and AddTime<=#" & date() &" 23:59:59#"
			Else
				NewsSql = "Select Count(ID) from FS_DownLoad where AuditTF=1 and AddTime>='" & date() &" 00:00:00' and AddTime<='" & date() &" 23:59:59'"
			End If
			Set RsNewsObj = Conn.Execute(NewsSql)
			TodayDownLoadNum = RsNewsObj(0)
			Set RsNewsObj = Nothing
		elseif RefreshType = "News" then
			ClassSql = "Select ClassID from FS_News where ClassID='" & RefreshID & "' order by ID desc"
			Set RsClassObj = Conn.Execute(ClassSql)
			if Not RsClassObj.Eof then
				TempClassID = ChildClassIDList(RsClassObj("ClassID"))
				AllClassID = "'" & RsClassObj("ClassID") & "'" & TempClassID
				TempClassIDArray = Split(AllClassID,",")
				ClassNum = UBound(TempClassIDArray)
				NewsSql = "Select Count(ID) from FS_DownLoad where AuditTF=1 and ClassID in (" & AllClassID & ")"
				Set RsNewsObj = Conn.Execute(NewsSql)
				NewsNum = RsNewsObj(0)
				Set RsNewsObj = Nothing
				NewsSql = "Select Sum(ClickNum) from FS_DownLoad where AuditTF=1 and ClassID in (" & AllClassID & ")"
				Set RsNewsObj = Conn.Execute(NewsSql)
				DownLoadClickNum = RsNewsObj(0)
				Set RsNewsObj = Nothing
				If IsSqlDataBase=0 then
					NewsSql = "Select Count(ID) from FS_DownLoad where AuditTF=1 and AddTime>=#" & date() &" 00:00:00# and AddTime<=#" & date() &" 23:59:59# and ClassID in (" & AllClassID & ")"
				Else
					NewsSql = "Select Count(ID) from FS_DownLoad where AuditTF=1 and AddTime>='" & date() &" 00:00:00' and AddTime<='" & date() &" 23:59:59' and ClassID in (" & AllClassID & ")"
				End If
				Set RsNewsObj = Conn.Execute(NewsSql)
				TodayDownLoadNum = RsNewsObj(0)
				Set RsNewsObj = Nothing
			else
				ClassNum = 0
				NewsNum = 0
				DownLoadClickNum = 0
				TodayDownLoadNum = 0
			end if
			Set RsClassObj = Nothing
		elseif RefreshType = "DownLoad" then
			ClassSql = "Select ClassID from FS_DownLoad where ClassID='" & RefreshID & "' order by ID desc"
			Set RsClassObj = Conn.Execute(ClassSql)
			if Not RsClassObj.Eof then
				TempClassID = ChildClassIDList(RsClassObj("ClassID"))
				AllClassID = "'" & RsClassObj("ClassID") & "'" & TempClassID
				TempClassIDArray = Split(AllClassID,",")
				ClassNum = UBound(TempClassIDArray)
				NewsSql = "Select Count(ID) from FS_DownLoad where AuditTF=1 and ClassID in (" & AllClassID & ")"
				Set RsNewsObj = Conn.Execute(NewsSql)
				NewsNum = RsNewsObj(0)
				Set RsNewsObj = Nothing
				NewsSql = "Select Sum(ClickNum) from FS_DownLoad where AuditTF=1 and ClassID in (" & AllClassID & ")"
				Set RsNewsObj = Conn.Execute(NewsSql)
				DownLoadClickNum = RsNewsObj(0)
				Set RsNewsObj = Nothing
				If IsSqlDataBase=0 then
					NewsSql = "Select Count(ID) from FS_DownLoad where AuditTF=1 and AddTime>=#" & date() &" 00:00:00# and AddTime<=#" & date() &" 23:59:59# and ClassID in (" & AllClassID & ")"
				Else
					NewsSql = "Select Count(ID) from FS_DownLoad where AuditTF=1 and AddTime>='" & date() &" 00:00:00' and AddTime<='" & date() &" 23:59:59' and ClassID in (" & AllClassID & ")"
				End If
				Set RsNewsObj = Conn.Execute(NewsSql)
				TodayDownLoadNum = RsNewsObj(0)
				Set RsNewsObj = Nothing
			else
				ClassNum = 0
				NewsNum = 0
				DownLoadClickNum = 0
				TodayDownLoadNum = 0
			end if
			Set RsClassObj = Nothing
		else
			ClassSql = ""
		end if
	else
		ClassSql = "Select ClassID from FS_NewsClass where ClassEName='" & ClassListStr & "' order by AddTime desc"
		Set RsClassObj = Conn.Execute(ClassSql)
		if Not RsClassObj.Eof then
			TempClassID = ChildClassIDList(RsClassObj("ClassID"))
			AllClassID = "'" & RsClassObj("ClassID") & "'" & TempClassID
			TempClassIDArray = Split(AllClassID,",")
			ClassNum = UBound(TempClassIDArray)
			NewsSql = "Select Count(ID) from FS_DownLoad where AuditTF=1 and ClassID in (" & AllClassID & ")"
			Set RsNewsObj = Conn.Execute(NewsSql)
			NewsNum = RsNewsObj(0)
			Set RsNewsObj = Nothing
			NewsSql = "Select Sum(ClickNum) from FS_DownLoad where AuditTF=1 and ClassID in (" & AllClassID & ")"
			Set RsNewsObj = Conn.Execute(NewsSql)
			DownLoadClickNum = RsNewsObj(0)
			Set RsNewsObj = Nothing
			If IsSqlDateBase=0 then
				NewsSql = "Select Count(ID) from FS_DownLoad where AuditTF=1 and AddTime>=#" & date() &" 00:00:00# and AddTime<=#" & date() &" 23:59:59# and ClassID in (" & AllClassID & ")"
			Else
				NewsSql = "Select Count(ID) from FS_DownLoad where AuditTF=1 and AddTime>='" & date() &" 00:00:00' and AddTime<='" & date() &" 23:59:59' and ClassID in (" & AllClassID & ")"
			End If
			Set RsNewsObj = Conn.Execute(NewsSql)
			TodayDownLoadNum = RsNewsObj(0)
			Set RsNewsObj = Nothing
		else
			ClassNum = 0
			NewsNum = 0
			DownLoadClickNum = 0
			TodayDownLoadNum = 0
		end if
		Set RsClassObj = Nothing
	end if
	UserSql = "Select Count(ID) from FS_Members"
	Set RsUserObj = Conn.Execute(UserSql)
	UserNum = RsUserObj(0)
	Set RsUserObj = Nothing  
	DownInfoStat = "<table Class=""" & CssFileStr & """ cellpadding=""0"" cellspacing=""0"" border=""0"" width=""100%"">" & Chr(13) & Chr(10)
	if ShowModeStr = "1" then
		DownInfoStat = DownInfoStat & "<tr>" & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "<td>" & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "栏目数量&nbsp;&nbsp;" & ClassNum  & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "</td>" & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "<td>" & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "下载数量&nbsp;&nbsp;" & NewsNum  & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "</td>" & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "<td>" & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "下载次数&nbsp;&nbsp;" & DownLoadClickNum  & Chr(13) & Chr(10) 
		DownInfoStat = DownInfoStat & "</td>" & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "<td>" & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "下载文件空间&nbsp;&nbsp;" & DownLoadFolderSize  & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "</td>" & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "<td>" & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "今日更新&nbsp;&nbsp;" & TodayDownLoadNum  & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "</td>" & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "<td>" & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "会员数量&nbsp;&nbsp;" & UserNum  & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "</td>" & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "</tr>" & Chr(13) & Chr(10)
	else
		DownInfoStat = DownInfoStat & "<tr>" & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "<td>" & "栏目数量&nbsp;&nbsp;" & ClassNum & "</td>" & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "</tr>" & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "<tr>" & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "<td>" & "下载数量&nbsp;&nbsp;" & NewsNum & "</td>" & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "</tr>" & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "<tr>" & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "<td>" & "下载次数&nbsp;&nbsp;" & DownLoadClickNum & "</td>" & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "</tr>" & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "<tr>" & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "<td>" & "下载文件空间&nbsp;&nbsp;" & DownLoadFolderSize & "</td>" & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "</tr>" & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "<tr>" & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "<td>" & "今日更新&nbsp;&nbsp;" & TodayDownLoadNum & "</td>" & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "</tr>" & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "<tr>" & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "<td>" & "会员数量&nbsp;&nbsp;" & UserNum & "</td>" & Chr(13) & Chr(10)
		DownInfoStat = DownInfoStat & "</tr>" & Chr(13) & Chr(10)
	end if
	DownInfoStat = DownInfoStat & "</table>"
End Function

Function DownLoad_Pic(PicWidth,PicHeight)
	DownLoad_Pic = "<img border=""0"" width="""& PicWidth & """ height=""" & PicHeight & """ src=""{DownLoad_Pic}"">"
End Function

Function AllChildClassIDStrList(ClassID)
	Dim TempRs,AllChildClassIDStrListA
	Set TempRs = Conn.Execute("Select ClassID,ChildNum from FS_NewsClass where ParentID = '" & ClassID & "' and DelFlag=0 order by AddTime desc ")
	do while Not TempRs.Eof
		AllChildClassIDStrList = AllChildClassIDStrList & ",'" & TempRs("ClassID") & "'"
		AllChildClassIDStrList = AllChildClassIDStrList & AllChildClassIDStrList(TempRs("ClassID"))
		TempRs.MoveNext
	loop

	TempRs.Close
	Set TempRs = Nothing
End Function
'读取数据生成自由标签
Function FreeLable(FreeLableID,QueryNum,ColSpan,RowSpan,ColNum,RowNum)
	Dim SqlStr,Rs,StyleContent,FieldsName,StartFlag,EndFlag
	FreeLableID = Replace(FreeLableID,"'","")
	QueryNum = Replace(Replace(QueryNum,"'","")," ","")
	SqlStr = "Select Sql,StyleContent,StartFlag,EndFlag From FS_FreeLable where FreeLableID = '"&FreeLableID&"'"
	Set Rs = conn.Execute(SqlStr)
	If Not Rs.eof Then
		'读取标签数据
		SqlStr = Replace(Rs("Sql"),"*|*","'")
		StyleContent = Replace(Rs("StyleContent"),"*|*","'")
		StartFlag = Rs("StartFlag")
		EndFlag = Rs("EndFlag")
		'生成自由标签
		FreeLable = CreateFreeLable(SqlStr,StyleContent,QueryNum,StartFlag,EndFlag,ColSpan,RowSpan,ColNum,RowNum)
	Else
		FreeLable = "无效的自由标签"
	End if
End Function
'生成自由标签
Function CreateFreeLable(SqlStr,StyleContent,QueryNum,StartFlag,EndFlag,ColSpan,RowSpan,ColNum,RowNum)
		Dim Rs,FieldsNameStr,FieldsNameArray,ExpressionArray,DateArray,PreDefineArray,ColSpanStr,RowSpanStr,TempRegExp,TempMatches,RecordIndex
		Dim RepeatContent,RepeatNumPerUnit,regEx,Match,Matches,TempMatch,indexOfField,i,j,k,l,m,TempExpression,TempResult,TempDate,DateFormtDictionary
		Dim TempRepeatContent,TempStr,TempFieldContent,TempStyleContent,NotRepeatFlag,RepeatFlag,HadShowFlag
		Dim NotRepeatContent
		Dim NotRepeatIndex
		NotRepeatFlag = false
		RepeatFlag = false
		TempStyleContent = StyleContent
		'从SQL语句中分解出字段
		TempStr = SqlStr
		If InStr(TempStr,"Top") <> 0 Then
			TempStr = Mid(TempStr,InStr(InStr(TempStr,"Top ")+4,TempStr," ")+1)
		Else
			TempStr = Mid(TempStr,8)
		End if
		FieldsNameStr = Left(TempStr,InStr(TempStr," from")-1)
		FieldsNameArray = split(Trim(FieldsNameStr),",")
				
		'设置查询数量,默认为10
		If QueryNum <> "" Then
			If InStr(SqlStr,"Top ") <> 0 Then
				i = InStr(SqlStr,"Top ") + 4
				j = InStr(i,SqlStr," ")
				TempStr = Mid(SqlStr,i,j-i)
				SqlStr = Replace(SqlStr,"Top "&TempStr,"Top "&QueryNum)
			else
				SqlStr = Replace(SqlStr,"Select","Select Top "&QueryNum)
			End if
			'Response.Write(QueryNum&"|"&TempStr&"|"&SqlStr)
		Else
			If InStr(SqlStr,"Top ") = 0 Then
				QueryNum = "10"
				SqlStr = Replace(SqlStr,"Select","Select Top "&QueryNum)
			else
				i = InStr(SqlStr,"Top ") + 4
				j = InStr(i,SqlStr," ")
				QueryNum = Mid(SqlStr,i,j-i)
			End if
		End if
		
		'计算重复数
		i = eval(ColNum&" * "&RowNum)
		j = eval(QueryNum)
		if j/i > int(j/i) Then
			RepeatNumPerUnit = int(j/i) + 1
		else
			RepeatNumPerUnit = int(j/i)
		end if
		
		'提取不重复样式内容
		set regEx = New RegExp
		regEx.Pattern = "{\*[1-9]+[0-9]*(\{[^\*]|[^\{]\*[^\}]|[^\*]\}|[^\{\*\}])*\*\}"
		regEx.IgnoreCase = True
		regEx.Global = True
		set TempRegExp = New RegExp
		TempRegExp.Pattern = "{\*[1-9]+[0-9]*"
		TempRegExp.IgnoreCase = True
		TempRegExp.Global = True
		Set Matches = regEx.Execute(TempStyleContent)
		If Matches.count > 0 Then
			ReDim NotRepeatContent(Matches.count)
			ReDim NotRepeatIndex(Matches.count)
			For	i = 0 to Matches.count-1
				NotRepeatContent(i) = Matches.item(i)
				Set TempMatches = TempRegExp.Execute(NotRepeatContent(i))
				If TempMatches.count > 0 Then
					NotRepeatIndex(i) = Cint(Mid(TempMatches.item(0),3))
				else
					NotRepeatIndex(i) = 0
				end if
			Next
			NotRepeatFlag = true
		End if
		
		'提取重复样式内容
		regEx.Pattern = "{#({[^#]|[^{]#[^}]|[^#]}|[^{#}])*#}"
		Set Matches = regEx.Execute(TempStyleContent)
		If Matches.count > 0 Then
			RepeatContent = Matches.item(0)
			RepeatFlag = true
		End if
		'如果没有不重复内容和重复内容就指定全部为不重复内容
		If not (NotRepeatFlag Or RepeatFlag) Then
			ReDim NotRepeatContent(0)
			ReDim NotRepeatIndex(0)
			NotRepeatContent(0) = TempStyleContent
			NotRepeatIndex(0) = 0
			NotRepeatFlag = true
		End if

		'从重复样式内容中提取表达式
		regEx.Pattern = "\(#(\([^#]|[^\(]#[^\)]|[^#]\)|[^\(#\)])*#\)"
		Set Matches = regEx.Execute(TempStyleContent)
		i = 0
		ReDim ExpressionArray(Matches.count)
		For each Match In Matches
			TempStr = Match
			TempStr = Replace(TempStr,"(#","")
			TempStr = Replace(TempStr,"#)","")
			ExpressionArray(i) = TempStr
			i = i + 1
		Next
		
		'从重复样式内容中提取日期样式
		regEx.Pattern = "\(\$(\([^\$]|[^\(]\$[^\)]|[^\$]\)|[^\(\$\)])*\$\)"
		Set Matches = regEx.Execute(TempStyleContent)
		i = 0
		ReDim DateArray(Matches.count)
		For each Match In Matches
			TempStr = Match
			TempStr = Replace(TempStr,"($","")
			TempStr = Replace(TempStr,"$)","")
			DateArray(i) = TempStr
			i = i + 1
		Next
		
		'从重复样式内容中提取系统预定义内容
		regEx.Pattern = "\[#(\[[^#]|[^\[]#[^\]]|[^#]\]|[^\[\]#])*#\]"
		Set Matches = regEx.Execute(TempStyleContent)
		i = 0
		ReDim PreDefineArray(Matches.count)
		For each Match In Matches
			TempStr = Match
			TempStr = Replace(TempStr,"[#","")
			TempStr = Replace(TempStr,"#]","")
			PreDefineArray(i) = TempStr
			i = i + 1
		Next
		
		Set Rs = conn.Execute(SqlStr)
		if RowSpan <> "" Then
			RowSpanStr = " height='"&RowSpan&"'"
		End if
		if ColSpan <> "" Then
			ColSpanStr = " width='"&ColSpan&"'"
		End if
		
		RecordIndex = 0
		'生成不重复内容
		If 	NotRepeatFlag = true Then
			If not Rs.eof Then
				While not Rs.eof
					For l = 0 to UBound(NotRepeatContent)
						If NotRepeatIndex(l) - 1 = RecordIndex Then
							TempRepeatContent = Replace(Replace(Replace(Replace(NotRepeatContent(l),"{*"&NotRepeatIndex(l),""),"{*0",""),"{*",""),"*}","")
							TempRepeatContent = CreateContentByRs(Rs,TempRepeatContent,FieldsNameArray,ExpressionArray,DateArray,PreDefineArray)
							TempStyleContent = Replace(TempStyleContent,NotRepeatContent(l),TempRepeatContent)
						End if
					Next
					RecordIndex = RecordIndex + 1
					Rs.MoveNext
				Wend
				Rs.movefirst
			End if
			'清除无效记录序号的不重复内容
			For l = 0 to UBound(NotRepeatContent)
				TempStyleContent = Replace(TempStyleContent,NotRepeatContent(l),"")
			Next
		End if

		'response.write(TempStyleContent)
		'response.end
		'按行、列、重复数生成标签
		RecordIndex = 0
		If not Rs.eof Then
			CreateFreeLable = StartFlag
			CreateFreeLable = CreateFreeLable&"<table>"
			For i = 1 to eval(RowNum)
				CreateFreeLable = CreateFreeLable & "<tr>"
				For j = 1 to eval(ColNum)
					CreateFreeLable = CreateFreeLable & "<td valign=top"&RowSpanStr&ColSpanStr&">"
					TempStr = ""
					'生成重复内容
					For l = 1 to RepeatNumPerUnit
						If 	NotRepeatFlag = true and Not Rs.eof Then
							Do
								HadShowFlag = false
								For m = 0 to UBound(NotRepeatIndex)
									if NotRepeatIndex(m) - 1 = RecordIndex Then
										HadShowFlag = true
										Exit For
									End if
								Next
								if HadShowFlag = true Then
									RecordIndex = RecordIndex +1
									Rs.MoveNext
								End if
							Loop While not Rs.eof and HadShowFlag = true
						End if
						If RepeatFlag = true Then
							If Not Rs.eof Then
								TempRepeatContent = Replace(Replace(RepeatContent,"{#",""),"#}","")
								TempRepeatContent = CreateContentByRs(Rs,TempRepeatContent,FieldsNameArray,ExpressionArray,DateArray,PreDefineArray)
								TempStr = TempStr&TempRepeatContent
								RecordIndex = RecordIndex +1
								Rs.movenext
							Else
								Exit For
							End if
						End if
					Next
					if err <> 0 Then
						err.clear()
					End if
					CreateFreeLable = CreateFreeLable&Replace(TempStyleContent,RepeatContent,TempStr)
					CreateFreeLable = CreateFreeLable & "</td>"
				Next
				CreateFreeLable = CreateFreeLable & "</tr>"
			Next
			CreateFreeLable = CreateFreeLable & "</table>"
			CreateFreeLable = CreateFreeLable & EndFlag
		End if
		Set ExpressionArray = nothing
		Set FieldsNameArray = nothing
		Set Rs = nothing
End Function
Function CreateContentByRs(Rs,TempRepeatContent,FieldsNameArray,ExpressionArray,DateArray,PreDefineArray)
	Dim TempExpression,Matches,Match,TempMatches,TempMatch,TempFieldContent,k,m,TempResult,TempDate,DateFormtDictionary,regEx
	'生成日期格式字典		
	Set DateFormtDictionary = CreateObject("Scripting.Dictionary")
	DateFormtDictionary.add "yyyy-mm-dd","1"
	DateFormtDictionary.add "yyyy.mm.dd","2"
	DateFormtDictionary.add "yyyy/mm/dd","3"
	DateFormtDictionary.add "mm/dd/yyyy","4"
	DateFormtDictionary.add "dd/mm/yyyy","5"
	DateFormtDictionary.add "mm-dd-yyyy","6"
	DateFormtDictionary.add "mm.dd.yyyy","7"
	DateFormtDictionary.add "mm-dd","8"
	DateFormtDictionary.add "mm/dd","9"
	DateFormtDictionary.add "mm.dd","10"
	DateFormtDictionary.add "mm月dd日","11"
	DateFormtDictionary.add "dd日hh时","12"
	DateFormtDictionary.add "dd日hh点","13"
	DateFormtDictionary.add "hh时mm分","14"
	DateFormtDictionary.add "hh:mm","15"
	DateFormtDictionary.add "yyyy年mm月dd日","16"
	set regEx = New RegExp
	regEx.IgnoreCase = True
	regEx.Global = True
	'生成表达式内容
	For k = 0 to UBound(ExpressionArray)
		TempExpression = ExpressionArray(k)
		regEx.Pattern = "[ ]*Left\([ ]*\[\*[^\[\*]*\*\][ ]*,[ ]*[1-9][0-9]*[ ]*\)"
		Set Matches = regEx.Execute(TempExpression)
		For each Match In Matches
			for m = 0 to UBound(FieldsNameArray)
				If InStr(Match,"[*"&FieldsNameArray(m)&"*]") <> 0 Then
					TempFieldContent = Rs(m)
					regEx.Pattern = "<\/*[^<>]*>"
					Set TempMatches = regEx.Execute(TempFieldContent)
					For each TempMatch In TempMatches
						TempFieldContent = Replace(TempFieldContent,TempMatch,"")
					Next
					TempExpression = Replace(TempExpression,Match,GotTopic(TempFieldContent,Clng(Mid(Match,InStrRev(Match,",")+1,InStrRev(Match,")")-InStrRev(Match,",")-1))))
					Exit for
				End if
			next
		Next
		TempRepeatContent = Replace(TempRepeatContent,"(#"&ExpressionArray(k)&"#)",TempExpression)
	Next
	
	'生成日期样式内容
	For k = 0 to UBound(DateArray)
		TempResult = ""
		TempDate = Trim(Lcase(DateArray(k)))
		for m = 0 to UBound(FieldsNameArray)
			If InStr(Lcase(Rs(m).name),"adddate") <> 0 Then
				TempResult = DateFormat(Rs(m),DateFormtDictionary(TempDate))
				Exit for
			End if
		next
		TempRepeatContent = Replace(TempRepeatContent,"($"&DateArray(k)&"$)",TempResult)
	Next
	'生成系统预定义内容
	For k = 0 to UBound(PreDefineArray)
		TempResult = ""
		Select Case Trim(Lcase(PreDefineArray(k)))
			case "url"
				for m = 0 to UBound(FieldsNameArray)
					If InStr(Lcase(Rs(m).name),"newsid") <> 0 Then
						TempResult = GetOneNewsLinkURL(Trim(Rs(m)))
						Exit for
					End if
					If InStr(Lcase(Rs(m).name),"downloadid") <> 0 Then
						TempResult = GetOneDownLoadLinkURL(Trim(Rs(m)))
						Exit for
					End if
				next
			case "classurl"
				for m = 0 to UBound(FieldsNameArray)
					If InStr(Lcase(Rs(m).name),"classid") <> 0 Then
						TempResult = GetOneClassLinkURLByID(Trim(Rs(m)))
						Exit for
					End if
				next
			case "picurl"
				for m = 0 to UBound(FieldsNameArray)
					If InStr(Lcase(Rs(m).name),"picpath") <> 0 Then
						GetAvailableDoMain
						TempResult = AvailableDoMain&(Trim(Rs(m)))
						Exit for
					End if
				next
		End Select
		TempRepeatContent = Replace(TempRepeatContent,"[#"&PreDefineArray(k)&"#]",TempResult)
	Next
	'生成字段内容

	For k = 0 to UBound(FieldsNameArray)
		TempFieldContent = Rs(k)
		if IsNull(TempFieldContent) Then
			TempFieldContent = ""
		End if
		TempRepeatContent = Replace(TempRepeatContent,"[*"&FieldsNameArray(k)&"*]",TempFieldContent)
	Next
								
	'清除多余字段							
	regEx.Pattern = "\[\*[^\[\]\*]*\*\]"
	Set Matches = regEx.Execute(TempRepeatContent)
	For each Match in Matches
		TempRepeatContent = Replace(TempRepeatContent,Match,"")
	Next
	CreateContentByRs = TempRepeatContent
End Function

'******************************
'根据ID得到当前所在栏目的名称
'author:lino
'Start
'*****************************
Function ypren()
Select Case RefreshType
  Case "Class"
   ypren = GetClassNameById(RefreshID)
  Case "News"
   ypren = GetNewsClassNameById(RefreshID)
  Case "Special"
   ypren = GetSpecialClassNameById(RefreshID)
  Case "DownLoad"
   ypren = GetDownloadClassNameById(RefreshID)
  Case Else
   ypren = ""
End Select
End Function

'栏目名称
Function GetClassNameById(ClassID)
Dim SqlClass,RsClassObj
if ClassID = "" then Exit Function

'**********如果是3.1版，把下行NewsClass改成FS_NewsClass
  Set RsClassObj = Conn.Execute("Select ClassCName from FS_NewsClass where ClassID='" & ClassID & "'")
  if Not RsClassObj.Eof then
   GetClassNameById=RsClassObj("ClassCName")
  end if 
RsClassObj.Close
Set RsClassObj = Nothing
End Function

'新闻所在的类目

Function GetNewsClassNameById(NewsID)
Dim SqlClass,RsClassObj
if NewsID = "" then Exit Function

'**********如果是3.1版，把下行News改成FS_News
  Set RsClassObj = Conn.Execute("Select ClassID from FS_News where NewsID='" & NewsID & "' ")
  if Not RsClassObj.Eof then
   GetNewsClassNameById=RsClassObj("ClassID")
   GetNewsClassNameById=GetClassNameById(GetNewsClassNameById)
  end if 
RsClassObj.Close
Set RsClassObj = Nothing
End Function

'专题所在的类目

Function GetSpecialClassNameById(SpecialID)
Dim SqlClass,RsClassObj
if SpecialID = "" then Exit Function

'**********如果是3.1版，把下行Special改成FS_Special
  Set RsClassObj = Conn.Execute("Select CName from FS_Special where SpecialID='" & SpecialID & "' ")
  if Not RsClassObj.Eof then
   GetSpecialClassNameById=RsClassObj("CName")
  end if 
RsClassObj.Close
Set RsClassObj = Nothing
End Function

'下载所在的类目

Function GetDownloadClassNameById(DownloadID)
Dim SqlClass,RsClassObj
if DownloadID = "" then Exit Function

'**********如果是3.1版，把下行Download改成FS_Download
Set RsClassObj = Conn.Execute("Select ClassID from FS_Download where DownloadID='" & DownloadID & "' ")
  if Not RsClassObj.Eof then
   GetDownloadClassNameById=RsClassObj("ClassID")
   GetDownloadClassNameById=GetClassNameById(GetDownloadClassNameById)
  end if 
RsClassObj.Close
Set RsClassObj = Nothing
End Function

'**************************
'End
'**************************

%>