<%
'==============================================================================
'软件名称：风讯网站信息管理系统
'当前版本：Foosun Content Manager System(FoosunCMS V3.1.0930)
'最新更新：2005.10
'==============================================================================
'Copyright (C) 2002-2004 Foosun.Net  All rights reserved.
'商业注册联系：028-85098980-601,项目开发：028-85098980-606、609,客户支持：608
'产品咨询QQ：394226379,159410,125114015
'技术支持QQ：315485710,66252421 
'项目开发QQ：415637671，655071
'程序开发：四川风讯科技发展有限公司(Foosun Inc.)
'Email:service@Foosun.cn
'MSN：skoolls@hotmail.com
'论坛支持：风讯在线论坛(http://bbs.foosun.net)
'官方网站：www.Foosun.cn  演示站点：test.cooin.com 
'网站通系列(智能快速建站系列)：www.ewebs.cn
'==============================================================================
'免费版本请在程序首页保留版权信息，并做上本站LOGO友情连接
'风讯公司保留此程序的法律追究权利
'如需进行2次开发，必须经过风讯公司书面允许。否则将追究法律责任
'==============================================================================
Dim JSCodeStr,i,TempClassObj,AvailableDoMain
Dim ListSpaces,ListSpaceStrs,Temp_ii
	ListSpaces = 3   '左右两条新闻之间的空格字符个数 
	ListSpaceStrs = ""
	for Temp_ii = 1 to ListSpaces
		ListSpaceStrs = ListSpaceStrs & "&nbsp;"
	next 
	
Dim RsSysJsConfig
Set RsSysJsConfig = Conn.Execute("Select DoMain from FS_Config")

Dim TempJSSysRootDir
if SysRootDir = "" then
	TempJSSysRootDir = ""
else
	TempJSSysRootDir = "/" & SysRootDir
end if

GetAvailableDoMain
Sub GetAvailableDoMain()
	Dim ConfigSql,RsConfigObj
	ConfigSql = "Select DoMain,MakeType,IndexExtName from FS_Config"
	Set RsConfigObj = Conn.Execute(ConfigSql)
	if Not RsConfigObj.Eof then
		AvailableDoMain = RsConfigObj("DoMain")
	else
		AvailableDoMain = GetDoMain
	end if
	Set RsConfigObj = Nothing
End Sub

Function CreateSysJS(FileName)'栏目JS新闻列表
	Dim RsSysJsObj,ClassIDStr,NewsNum,MarDirection,BrStr,MarSpeed,NaviPic,RsCreateSql,RsCreateObj,DateTF,RowNum,RowSpace,TitleNum,ShowClassTF
	Dim RightDate,ClassID,RsClassObj,PicHeight,MarWidth,OpenMode,MarHeight,PicWidth,ShowTitle,TitleCSS,SaveFilePath,FileNameStr,DateCSS,DateType,LinkCSS,MoreContentStr,MoreContentTF
	Set RsSysJsObj = Conn.Execute("Select * from FS_SysJs where FileName='"&FileName&"'")
	If Not RsSysJsObj.eof then
		ClassID = RsSysJsObj("ClassID")
		If RsSysJsObj("NaviPic")<>"" then
			NaviPic = "<img src=""" & AvailableDoMain & RsSysJsObj("NaviPic") & """ border=""0"">"
		Else
			NaviPic = ""
		End If
		NewsNum = RsSysJsObj("NewsNum")
		RowNum = RsSysJsObj("RowNum")
		RowSpace = RsSysJsObj("RowSpace")
		TitleNum = RsSysJsObj("TitleNum")
		TitleCSS = RsSysJsObj("TitleCSS")
		SaveFilePath = RsSysJsObj("FileSavePath")
		FileNameStr = RsSysJsObj("FileName")
		DateCSS = RsSysJsObj("DateCSS")
		DateType = RsSysJsObj("DateType")
		MarDirection = RsSysJsObj("MarDirection")
		MarSpeed = RsSysJsObj("MarSpeed")
		PicWidth = RsSysJsObj("PicWidth")
		PicHeight = RsSysJsObj("PicHeight")
		MarWidth = RsSysJsObj("MarWidth")
		MarHeight = RsSysJsObj("MarHeight")
		If RsSysJsObj("OpenMode")=1 then
			OpenMode = " target=""_blank"""
		Else
			OpenMode = " target=""_self"""
		End If
		If RsSysJsObj("ShowTitle")<>0 then
			ShowTitle = True
		Else
			ShowTitle = false
		End If
		If RsSysJsObj("MarDirection")="left" or RsSysJsObj("MarDirection")="right" then
			BrStr = ""
		Else
			BrStr = "<br>"
		End If
		If RsSysJsObj("MoreContent")<>0 then
			MoreContentTF = True
			MoreContentStr = RsSysJsObj("LinkWord")
			LinkCSS = RsSysJsObj("LinkCSS")
		Else
			MoreContentTF = False
		End If
		If RsSysJsObj("DateType")<>0 then
			DateTF = true
		Else
			DateTF = false
		End If
		If RsSysJsObj("ClassName")<>0 then
			ShowClassTF = true
		Else
			ShowClassTF = false
		End If
		If RsSysJsObj("RightDate")<>0 then
			RightDate = true
		Else
			RightDate = false
		End If
		ClassIDStr = ClassID
		If RsSysJsObj("SonClass")=1 then
			Set RsClassObj = Conn.Execute("Select ClassID from FS_NewsClass where ParentID='"&ClassID&"' and DelFlag=0 order by ID desc")
			If Not RsClassObj.eof then
				Do While not RsClassObj.eof
					ClassIDStr = ClassIDStr &"','"& RsClassObj("ClassID")
					RsClassObj.MoveNext
				Loop
				RsClassObj.Close
				Set RsClassObj = Nothing
			End If
		End If
		Select Case RsSysJsObj("NewsType")
			Case "RecNews"
				if RsSysJsObj("FileType")=1 then
					RsCreateSql = "Select top "&NewsNum&" * From FS_News where ClassID in ('"&ClassIDStr&"') and DelTF=0 and RecTF=1 and AuditTF=1 order by AddDate desc" '推荐新闻
				else
					RsCreateSql = "Select top "&NewsNum&" * From FS_News where DelTF=0 and RecTF=1 and AuditTF=1 order by AddDate desc" '推荐新闻
				end if
				Set RsCreateObj = Conn.Execute(RsCreateSql)
				If Not RsCreateObj.eof then
					JSCodeStr = "document.write('<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing="""&RowSpace&"""><tr>"
					for i=1 to NewsNum
					  If RsCreateObj.eof then Exit For
					  Set TempClassObj = Conn.Execute("Select ClassEName,ClassCName,SaveFilePath from FS_NewsClass where ClassID='"&RsCreateObj("ClassID")&"'")
					  If DateTF = true then
						If ShowClassTF = true then
							If RightDate = true then
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"["&TempClassObj("ClassCName")&"]"&"<a class="""&TitleCSS&""" href=""" & GetOneNewsLinkURL(RsCreateObj("NewsID")) &""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a></td><td><div align=Right><Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</div></td>"
							Else
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"["&TempClassObj("ClassCName")&"]"&"<a class="""&TitleCSS&""" href=""" & GetOneNewsLinkURL(RsCreateObj("NewsID")) &""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</td>"
							End If
						Else
							If RightDate = true then
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"<a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID")) &""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a></td><td><div align=Right><Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</div></td>"
							Else
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"<a class="""&TitleCSS&""" href=""" & GetOneNewsLinkURL(RsCreateObj("NewsID")) &""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</td>"
							End If
						End IF
					  Else
						If ShowClassTF = true then
							  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"["&TempClassObj("ClassCName")&"]"&"<a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;"&ListSpaceStrs&"</td>"
						Else
							  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"<a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;"&ListSpaceStrs&"</td>"
						End If
					  End If
					  RsCreateObj.MoveNext
					  if i mod Cint(RowNum) = 0 or RsCreateObj.eof then
						if RightDate = true then
						  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&RowNum*2&""" height=""1"" background=""" & AvailableDoMain & RsSysJsObj("RowBetween")&"""></td></tr><tr>"
						else
						  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&RowNum&""" height=""1"" background=""" & AvailableDoMain & RsSysJsObj("RowBetween")&"""></td></tr><tr>"
						end if
					  end if
					next 
					If RsSysJsObj("FileType")=1 then
						Dim RsTempClassObjs
					Set RsTempClassObjs = Conn.Execute("Select SaveFilePath,ClassEName,FileExtName from FS_NewsClass where ClassID='"&ClassID&"'")
						If RsTempClassObjs.eof then
							CreateSysJS = "刷新栏目已经不存在"
							Exit Function
						End If
					End If
					If RightDate = true then
						if RsSysJsObj("FileType")=1 and MoreContentTF=True then
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum*2&""" align=Right><a class="""&LinkCSS&""" href=""" & GetOneClassLinkURL(RsTempClassObjs("ClassEName"),RsTempClassObjs("SaveFilePath"),RsTempClassObjs("FileExtName")) & """>"& MoreContentStr&"</a>"&ListSpaceStrs&"</td></tr>"
						else
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum*2&""" height="&RowSpace&"></td></tr>"
						end if
					Else
						if RsSysJsObj("FileType")=1 and MoreContentTF=True then
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum&""" align=Right><a class="""&LinkCSS&""" href=""" & GetOneClassLinkURL(RsTempClassObjs("ClassEName"),RsTempClassObjs("SaveFilePath"),RsTempClassObjs("FileExtName")) & """>"& MoreContentStr&"</a>"&ListSpaceStrs&"</td></tr>"
						else
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum&""" height="&RowSpace&"></td></tr>"
						end if
					End If
					JSCodeStr = JSCodeStr & "</table>');"
					WriteFile SaveFilePath,FileNameStr,JSCodeStr '写文件
					Conn.Execute("Update FS_SysJs Set AddTime='"&Now()&"' where FileName='"&FileName&"'")
					CreateSysJS = true
					RsCreateObj.Close 
					Set RsCreateObj = Nothing 
				Else
					JSCodeStr = "document.write('未查询到符合条件的新闻')"
					WriteFile SaveFilePath,FileNameStr,JSCodeStr '写文件
					Conn.Execute("Update FS_SysJs Set AddTime='"&Now()&"' where FileName='"&FileName&"'")
					CreateSysJS = "文件添加或修改成功\n\n但未找到符合条件的新闻,建议您稍后再试"
				End If
			Case "MarqueeNews"
				if RsSysJsObj("FileType")=1 then
					RsCreateSql = "Select top "&NewsNum&" * From FS_News where ClassID in ('"&ClassIDStr&"') and DelTF=0 and MarqueeNews=1 and AuditTF=1 order by AddDate desc" '滚动新闻
				else
					RsCreateSql = "Select top "&NewsNum&" * From FS_News where DelTF=0 and MarqueeNews=1 and AuditTF=1 order by AddDate desc" '滚动新闻
				end if
				Set RsCreateObj = Conn.Execute(RsCreateSql)
				If Not RsCreateObj.eof then
					JSCodeStr = "document.write('<marquee onmouseout=start() onmouseover=stop() Width="&MarWidth&" Height="&MarHeight&" scrolldelay=80 direction="&MarDirection&" scrollamount="& CInt(MarSpeed) &">"
					for i=1 to NewsNum
					  If RsCreateObj.eof then Exit For
					  Set TempClassObj = Conn.Execute("Select ClassEName,ClassCName,SaveFilePath from FS_NewsClass where ClassID='"&RsCreateObj("ClassID")&"'")
					  If DateTF = true then
						If ShowClassTF = true then
							  JSCodeStr = JSCodeStr & NaviPic &"["&TempClassObj("ClassCName")&"]"&"<a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span>"&ListSpaceStrs&BrStr
						Else
							  JSCodeStr = JSCodeStr & NaviPic &"<a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span>"&ListSpaceStrs&BrStr
						End IF
					  Else
						If ShowClassTF = true then
							  JSCodeStr = JSCodeStr & NaviPic &"["&TempClassObj("ClassCName")&"]"&"<a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;"&ListSpaceStrs&BrStr
						Else
							  JSCodeStr = JSCodeStr & NaviPic &"<a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;"&ListSpaceStrs&BrStr
						End If
					  End If
					  RsCreateObj.MoveNext
					next 
					If RsSysJsObj("FileType")=1 then
					Set RsTempClassObjs = Conn.Execute("Select SaveFilePath,ClassEName,FileExtName from FS_NewsClass where ClassID='"&ClassID&"'")
						If RsTempClassObjs.eof then
							CreateSysJS = "刷新栏目已经不存在"
							Exit Function
						End If
					End If
					if RsSysJsObj("FileType")=1 and MoreContentTF=True then
						JSCodeStr = JSCodeStr &"<a class="""&LinkCSS&""" href=""" & GetOneClassLinkURL(RsTempClassObjs("ClassEName"),RsTempClassObjs("SaveFilePath"),RsTempClassObjs("FileExtName")) & """>"& MoreContentStr&"</a>"&ListSpaceStrs
					end if
					JSCodeStr = JSCodeStr & "</marquee>');"
					WriteFile SaveFilePath,FileNameStr,JSCodeStr '写文件
					Conn.Execute("Update FS_SysJs Set AddTime='"&Now()&"' where FileName='"&FileName&"'")
					CreateSysJS = true
					RsCreateObj.Close 
					Set RsCreateObj = Nothing 
				Else
					JSCodeStr = "document.write('未查询到符合条件的新闻')"
					WriteFile SaveFilePath,FileNameStr,JSCodeStr '写文件
					Conn.Execute("Update FS_SysJs Set AddTime='"&Now()&"' where FileName='"&FileName&"'")
					CreateSysJS = "文件添加或修改成功\n\n但未找到符合条件的新闻,建议您稍后再试"
				End If
			Case "SBSNews"
				if RsSysJsObj("FileType")=1 then
					RsCreateSql = "Select top "&NewsNum&" * From FS_News where ClassID in ('"&ClassIDStr&"') and DelTF=0 and SBSNews=1 and AuditTF=1 order by AddDate desc" '并排新闻
				else
					RsCreateSql = "Select top "&NewsNum&" * From FS_News where DelTF=0 and SBSNews=1 and AuditTF=1 order by AddDate desc" '并排新闻
				end if
				Set RsCreateObj = Conn.Execute(RsCreateSql)
				If Not RsCreateObj.eof then
					JSCodeStr = "document.write('<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing="""&RowSpace&"""><tr>"
					for i=1 to NewsNum
					  If RsCreateObj.eof then Exit For
					  Set TempClassObj = Conn.Execute("Select ClassEName,ClassCName,SaveFilePath from FS_NewsClass where ClassID='"&RsCreateObj("ClassID")&"'")
					  If DateTF = true then
						If ShowClassTF = true then
							If RightDate = true then
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"["&TempClassObj("ClassCName")&"]"&"<a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a></td><td><div align=Right><Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</div></td>"
							Else
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"["&TempClassObj("ClassCName")&"]"&"<a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</td>"
							End If
						Else
							If RightDate = true then
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"<a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a></td><td><div align=Right><Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</div></td>"
							Else
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"<a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</td>"
							End If
						End IF
					  Else
						If ShowClassTF = true then
							  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"["&TempClassObj("ClassCName")&"]"&"<a class="""&TitleCSS&""" href=""" & GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;"&ListSpaceStrs&"</td>"
						Else
							  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"<a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;"&ListSpaceStrs&"</td>"
						End If
					  End If
					  RsCreateObj.MoveNext
					  if i mod Cint(RowNum) = 0 or RsCreateObj.eof then
						if RightDate = true then
						  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&RowNum*2&""" height=""1"" background=""" & AvailableDoMain & RsSysJsObj("RowBetween")&"""></td></tr><tr>"
						else
						  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&RowNum&""" height=""1"" background=""" & AvailableDoMain & RsSysJsObj("RowBetween")&"""></td></tr><tr>"
						end if
					  end if
					next 
					If RsSysJsObj("FileType")=1 then
					Set RsTempClassObjs = Conn.Execute("Select SaveFilePath,ClassEName,FileExtName from FS_NewsClass where ClassID='"&ClassID&"'")
						If RsTempClassObjs.eof then
							CreateSysJS = "刷新栏目已经不存在"
							Exit Function
						End If
					End If
					If RightDate = true then
						if RsSysJsObj("FileType")=1 and MoreContentTF=True then
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum*2&""" align=Right><a class="""&LinkCSS&""" href="""& GetOneClassLinkURL(RsTempClassObjs("ClassEName"),RsTempClassObjs("SaveFilePath"),RsTempClassObjs("FileExtName")) & """>"& MoreContentStr&"</a>"&ListSpaceStrs&"</td></tr>"
						else
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum*2&""" height="&RowSpace&"></td></tr>"
						end if
					Else
						if RsSysJsObj("FileType")=1 and MoreContentTF=True then
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum&""" align=Right><a class="""&LinkCSS&""" href="""&GetOneClassLinkURL(RsTempClassObjs("ClassEName"),RsTempClassObjs("SaveFilePath"),RsTempClassObjs("FileExtName")) &""">"& MoreContentStr&"</a>"&ListSpaceStrs&"</td></tr>"
						else
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum&""" height="&RowSpace&"></td></tr>"
						end if
					End If
					JSCodeStr = JSCodeStr & "</table>');"
					WriteFile SaveFilePath,FileNameStr,JSCodeStr '写文件
					Conn.Execute("Update FS_SysJs Set AddTime='"&Now()&"' where FileName='"&FileName&"'")
					CreateSysJS = true
					RsCreateObj.Close 
					Set RsCreateObj = Nothing 
				Else
					JSCodeStr = "document.write('未查询到符合条件的新闻')"
					WriteFile SaveFilePath,FileNameStr,JSCodeStr '写文件
					Conn.Execute("Update FS_SysJs Set AddTime='"&Now()&"' where FileName='"&FileName&"'")
					CreateSysJS = "文件添加或修改成功\n\n但未找到符合条件的新闻,建议您稍后再试"
				End If
			Case "PicNews"
				if RsSysJsObj("FileType")=1 then
					RsCreateSql = "Select top "&NewsNum&" * From FS_News where ClassID in ('"&ClassIDStr&"') and DelTF=0 and PicNewsTF=1 and AuditTF=1 order by AddDate desc" '图片新闻
				else
					RsCreateSql = "Select top "&NewsNum&" * From FS_News where DelTF=0 and PicNewsTF=1 and AuditTF=1 order by AddDate desc" '图片新闻
				end if
				Set RsCreateObj = Conn.Execute(RsCreateSql)
				If Not RsCreateObj.eof then
					JSCodeStr = "document.write('<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing="""&RowSpace&"""><tr>"
					for i=1 to NewsNum
					  If RsCreateObj.eof then Exit For
					  Set TempClassObj = Conn.Execute("Select ClassEName,ClassCName,SaveFilePath from FS_NewsClass where ClassID='"&RsCreateObj("ClassID")&"'")
					  If DateTF = true then
						If ShowClassTF = true then
							If RightDate = true then
								If ShowTitle = True then
								  JSCodeStr = JSCodeStr &"<td align=center valign=middle><table border=0 cellspacing=0 cellpadding=0><tr><td colspan=2 align=center valign=middle><a href="&GetOneNewsLinkURL(RsCreateObj("NewsID"))&OpenMode&"><img src="&AvailableDoMain&RsCreateObj("PicPath")&" height="&PicHeight&" width="&PicWidth&" border=0></a></td></tr>"
								  JSCodeStr = JSCodeStr &"<tr><td align=center>"& NaviPic &"["&TempClassObj("ClassCName")&"]"&"<a class="""&TitleCSS&""" href=""" & GetOneNewsLinkURL(RsCreateObj("NewsID")) &""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a></td><td><div align=Right><Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span></div></td></tr></table></td>"
								Else
								  JSCodeStr = JSCodeStr &"<td align=center valign=middle><a href=" & GetOneNewsLinkURL(RsCreateObj("NewsID")) & OpenMode&"><img src="&AvailableDoMain&RsCreateObj("PicPath")&" height="&PicHeight&" width="&PicWidth&" border=0></a></td>"
								End If
							Else
								If ShowTitle = True then
								  JSCodeStr = JSCodeStr &"<td align=center valign=middle><table border=0 cellspacing=0 cellpadding=0><tr><td colspan=2 align=center valign=middle><a href="&GetOneNewsLinkURL(RsCreateObj("NewsID"))&OpenMode&"><img src="&AvailableDoMain&RsCreateObj("PicPath")&" height="&PicHeight&" width="&PicWidth&" border=0></a></td></tr>"
								  JSCodeStr = JSCodeStr &"<tr><td align=center>"& NaviPic &"["&TempClassObj("ClassCName")&"]"&"<a class="""&TitleCSS&""" href=""" & GetOneNewsLinkURL(RsCreateObj("NewsID")) &""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span></td></tr></table></td>"
								Else
								  JSCodeStr = JSCodeStr &"<td align=center valign=middle><a href=" & GetOneNewsLinkURL(RsCreateObj("NewsID")) & OpenMode&"><img src="&AvailableDoMain&RsCreateObj("PicPath")&" height="&PicHeight&" width="&PicWidth&" border=0></a></td>"
								End If
							End If
						Else
							If RightDate = true then
								If ShowTitle = True then
								  JSCodeStr = JSCodeStr &"<td align=center valign=middle><table border=0 cellspacing=0 cellpadding=0><tr><td colspan=2 align=center valign=middle><a href="&GetOneNewsLinkURL(RsCreateObj("NewsID"))&OpenMode&"><img src="&AvailableDoMain&RsCreateObj("PicPath")&" height="&PicHeight&" width="&PicWidth&" border=0></a></td></tr>"
								  JSCodeStr = JSCodeStr &"<tr><td align=center>"& NaviPic &"<a class="""&TitleCSS&""" href=""" & GetOneNewsLinkURL(RsCreateObj("NewsID")) &""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a></td><td><div align=Right><Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span></div></td></tr></table></td>"
								Else
								  JSCodeStr = JSCodeStr &"<td align=center valign=middle><a href=" & GetOneNewsLinkURL(RsCreateObj("NewsID")) & OpenMode&"><img src="&AvailableDoMain&RsCreateObj("PicPath")&" height="&PicHeight&" width="&PicWidth&" border=0></a></td>"
								End If
							Else
								If ShowTitle = True then
								  JSCodeStr = JSCodeStr &"<td align=center valign=middle><table border=0 cellspacing=0 cellpadding=0><tr><td colspan=2 align=center valign=middle><a href="&GetOneNewsLinkURL(RsCreateObj("NewsID"))&OpenMode&"><img src="&AvailableDoMain&RsCreateObj("PicPath")&" height="&PicHeight&" width="&PicWidth&" border=0></a></td></tr>"
								  JSCodeStr = JSCodeStr &"<tr><td align=center>"& NaviPic &"<a class="""&TitleCSS&""" href=""" & GetOneNewsLinkURL(RsCreateObj("NewsID")) &""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span></td></tr></table></td>"
								Else
								  JSCodeStr = JSCodeStr &"<td align=center valign=middle><a href=" & GetOneNewsLinkURL(RsCreateObj("NewsID")) & OpenMode & "><img src="&AvailableDoMain&RsCreateObj("PicPath")&" height="&PicHeight&" width="&PicWidth&" border=0></a></td>"
								End If
							End If
						End IF
					  Else
						If ShowClassTF = true then
							If ShowTitle = True then
							  JSCodeStr = JSCodeStr &"<td align=center valign=middle><table border=0 cellspacing=0 cellpadding=0><tr><td align=center valign=middle><a href=" & GetOneNewsLinkURL(RsCreateObj("NewsID")) & OpenMode & "><img src="&AvailableDoMain&RsCreateObj("PicPath")&" height="&PicHeight&" width="&PicWidth&" border=0></a></td></tr>"
							  JSCodeStr = JSCodeStr &"<tr><td align=center>"& NaviPic &"["&TempClassObj("ClassCName")&"]"&"<a class="""&TitleCSS&""" href=""" & GetOneNewsLinkURL(RsCreateObj("NewsID")) &""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a></td></tr></table></td>"
							Else
							  JSCodeStr = JSCodeStr &"<td align=center valign=middle><a href=" & GetOneNewsLinkURL(RsCreateObj("NewsID")) & OpenMode & "><img src="&AvailableDoMain&RsCreateObj("PicPath")&" height="&PicHeight&" width="&PicWidth&" border=0></a></td>"
							End If
						Else
							If ShowTitle = True then
							  JSCodeStr = JSCodeStr &"<td align=center valign=middle><table border=0 cellspacing=0 cellpadding=0><tr><td align=center valign=middle><a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&"><img src="&AvailableDoMain&RsCreateObj("PicPath")&" height="&PicHeight&" width="&PicWidth&" border=0 ></a></td></tr>"
							  JSCodeStr = JSCodeStr &"<tr><td align=center>"& NaviPic &"<a class="""&TitleCSS&""" href=""" & GetOneNewsLinkURL(RsCreateObj("NewsID")) & """" & OpenMode & ">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a></td></tr></table></td>"
							Else
							  JSCodeStr = JSCodeStr &"<td align=center valign=middle><a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&"><img src="&AvailableDoMain&RsCreateObj("PicPath")&" height="&PicHeight&" width="&PicWidth&" border=0 ></a></td>"
							End If
						End If
					  End If
					  RsCreateObj.MoveNext
					  if i mod Cint(RowNum) = 0 or RsCreateObj.eof then
						if RightDate = true then
						  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&RowNum*2&""" height=""1"" background=""" & AvailableDoMain & RsSysJsObj("RowBetween")&"""></td></tr><tr>"
						else
						  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&RowNum&""" height=""1"" background=""" & AvailableDoMain & RsSysJsObj("RowBetween")&"""></td></tr><tr>"
						end if
					  end if
					next 
					If RsSysJsObj("FileType")=1 then
					Set RsTempClassObjs = Conn.Execute("Select SaveFilePath,ClassEName,FileExtName from FS_NewsClass where ClassID='"&ClassID&"'")
						If RsTempClassObjs.eof then
							CreateSysJS = "刷新栏目已经不存在"
							Exit Function
						End If
					End If
					If RightDate = true then
						if RsSysJsObj("FileType")=1 and MoreContentTF=True then
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum*2&""" align=Right><a class="""&LinkCSS&""" href=""" & GetOneClassLinkURL(RsTempClassObjs("ClassEName"),RsTempClassObjs("SaveFilePath"),RsTempClassObjs("FileExtName")) & """>"& MoreContentStr&"</a>"&ListSpaceStrs&"</td></tr>"
						else
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum*2&""" height="&RowSpace&"></td></tr>"
						end if
					Else
						if RsSysJsObj("FileType")=1 and MoreContentTF=True then
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum&""" align=Right><a class="""&LinkCSS&""" href=""" & GetOneClassLinkURL(RsTempClassObjs("ClassEName"),RsTempClassObjs("SaveFilePath"),RsTempClassObjs("FileExtName")) & """>"& MoreContentStr&"</a>"&ListSpaceStrs&"</td></tr>"
						else
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum&""" height="&RowSpace&"></td></tr>"
						end if
					End If
					JSCodeStr = JSCodeStr & "</table>');"
					WriteFile SaveFilePath,FileNameStr,JSCodeStr '写文件
					Conn.Execute("Update FS_SysJs Set AddTime='"&Now()&"' where FileName='"&FileName&"'")
					CreateSysJS = true
					RsCreateObj.Close 
					Set RsCreateObj = Nothing 
				Else
					JSCodeStr = "document.write('未查询到符合条件的新闻')"
					WriteFile SaveFilePath,FileNameStr,JSCodeStr '写文件
					Conn.Execute("Update FS_SysJs Set AddTime='"&Now()&"' where FileName='"&FileName&"'")
					CreateSysJS = "文件添加或修改成功\n\n但未找到符合条件的新闻,建议您稍后再试"
				End If
			Case "NewNews"
				if RsSysJsObj("FileType")=1 then
					RsCreateSql = "Select top "&NewsNum&" * From FS_News where ClassID in ('"&ClassIDStr&"') and DelTF=0 and AuditTF=1 order by AddDate desc" '最新新闻
				else
					RsCreateSql = "Select top "&NewsNum&" * From FS_News where DelTF=0 and AuditTF=1 order by AddDate desc" '最新新闻
				end if
				Set RsCreateObj = Conn.Execute(RsCreateSql)
				If Not RsCreateObj.eof then
					JSCodeStr = "document.write('<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing="""&RowSpace&"""><tr>"
					for i=1 to NewsNum
					  If RsCreateObj.eof then Exit For
					  Set TempClassObj = Conn.Execute("Select ClassEName,ClassCName,SaveFilePath from FS_NewsClass where ClassID='"&RsCreateObj("ClassID")&"'")
					  If DateTF = true then
						If ShowClassTF = true then
							If RightDate = true then
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"["&TempClassObj("ClassCName")&"]"&"<a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a></td><td><div align=Right><Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</div></td>"
							Else
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"["&TempClassObj("ClassCName")&"]"&"<a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</td>"
							End If
						Else
							If RightDate = true then
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"<a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a></td><td><div align=Right><Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</div></td>"
							Else
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"<a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</td>"
							End If
						End IF
					  Else
						If ShowClassTF = true then
							  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"["&TempClassObj("ClassCName")&"]"&"<a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;"&ListSpaceStrs&"</td>"
						Else
							  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"<a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;"&ListSpaceStrs&"</td>"
						End If
					  End If
					  RsCreateObj.MoveNext
					  if i mod Cint(RowNum) = 0 or RsCreateObj.eof then
						if RightDate = true then
						  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&RowNum*2&""" height=""1"" background=""" & AvailableDoMain & RsSysJsObj("RowBetween")&"""></td></tr><tr>"
						else
						  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&RowNum&""" height=""1"" background=""" & AvailableDoMain & RsSysJsObj("RowBetween")&"""></td></tr><tr>"
						end if
					  end if
					next 
					If RsSysJsObj("FileType")=1 then
					Set RsTempClassObjs = Conn.Execute("Select SaveFilePath,ClassEName,FileExtName from FS_NewsClass where ClassID='"&ClassID&"'")
						If RsTempClassObjs.eof then
							CreateSysJS = "刷新栏目已经不存在"
							Exit Function
						End If
					End If
					If RightDate = true then
						if RsSysJsObj("FileType")=1 and MoreContentTF=True then
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum*2&""" align=Right><a class="""&LinkCSS&""" href="""&GetOneClassLinkURL(RsTempClassObjs("ClassEName"),RsTempClassObjs("SaveFilePath"),RsTempClassObjs("FileExtName")) & """>"& MoreContentStr&"</a>"&ListSpaceStrs&"</td></tr>"
						else
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum*2&""" height="&RowSpace&"></td></tr>"
						end if
					Else
						if RsSysJsObj("FileType")=1 and MoreContentTF=True then
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum&""" align=Right><a class="""&LinkCSS&""" href=""" & GetOneClassLinkURL(RsTempClassObjs("ClassEName"),RsTempClassObjs("SaveFilePath"),RsTempClassObjs("FileExtName")) & """>"& MoreContentStr&"</a>"&ListSpaceStrs&"</td></tr>"
						else
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum&""" height="&RowSpace&"></td></tr>"
						end if
					End If
					JSCodeStr = JSCodeStr & "</table>');"
					WriteFile SaveFilePath,FileNameStr,JSCodeStr '写文件
					Conn.Execute("Update FS_SysJs Set AddTime='"&Now()&"' where FileName='"&FileName&"'")
					CreateSysJS = true
					RsCreateObj.Close 
					Set RsCreateObj = Nothing 
				Else
					JSCodeStr = "document.write('未查询到符合条件的新闻')"
					WriteFile SaveFilePath,FileNameStr,JSCodeStr '写文件
					Conn.Execute("Update FS_SysJs Set AddTime='"&Now()&"' where FileName='"&FileName&"'")
					CreateSysJS = "文件添加或修改成功\n\n但未找到符合条件的新闻,建议您稍后再试"
				End If
			Case "HotNews"
				if RsSysJsObj("FileType")=1 then
					RsCreateSql = "Select top "&NewsNum&" * From FS_News where ClassID in ('"&ClassIDStr&"') and DelTF=0 and AuditTF=1 order by ClickNum desc" '热点新闻
				else
					RsCreateSql = "Select top "&NewsNum&" * From FS_News where DelTF=0 and AuditTF=1 order by ClickNum desc" '热点新闻
				end if
				Set RsCreateObj = Conn.Execute(RsCreateSql)
				If Not RsCreateObj.eof then
					JSCodeStr = "document.write('<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing="""&RowSpace&"""><tr>"
					for i=1 to NewsNum
					  If RsCreateObj.eof then Exit For
					  Set TempClassObj = Conn.Execute("Select ClassEName,ClassCName,SaveFilePath from FS_NewsClass where ClassID='"&RsCreateObj("ClassID")&"'")
					  If DateTF = true then
						If ShowClassTF = true then
							If RightDate = true then
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"["&TempClassObj("ClassCName")&"]"&"<a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a></td><td><div align=Right><Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</div></td>"
							Else
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"["&TempClassObj("ClassCName")&"]"&"<a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</td>"
							End If
						Else
							If RightDate = true then
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"<a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a></td><td><div align=Right><Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</div></td>"
							Else
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"<a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</td>"
							End If
						End IF
					  Else
						If ShowClassTF = true then
							  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"["&TempClassObj("ClassCName")&"]"&"<a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;"&ListSpaceStrs&"</td>"
						Else
							  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"<a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;"&ListSpaceStrs&"</td>"
						End If
					  End If
					  RsCreateObj.MoveNext
					  if i mod Cint(RowNum) = 0 or RsCreateObj.eof then
						if RightDate = true then
						  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&RowNum*2&""" height=""1"" background=""" & AvailableDoMain & RsSysJsObj("RowBetween")&"""></td></tr><tr>"
						else
						  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&RowNum&""" height=""1"" background=""" & AvailableDoMain & RsSysJsObj("RowBetween")&"""></td></tr><tr>"
						end if
					  end if
					next 
					If RsSysJsObj("FileType")=1 then
					Set RsTempClassObjs = Conn.Execute("Select SaveFilePath,ClassEName,FileExtName from FS_NewsClass where ClassID='"&ClassID&"'")
						If RsTempClassObjs.eof then
							CreateSysJS = "刷新栏目已经不存在"
							Exit Function
						End If
					End If
					If RightDate = true then
						if RsSysJsObj("FileType")=1 and MoreContentTF=True then
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum*2&""" align=Right><a class="""&LinkCSS&""" href=""" & GetOneClassLinkURL(RsTempClassObjs("ClassEName"),RsTempClassObjs("SaveFilePath"),RsTempClassObjs("FileExtName")) &""">"& MoreContentStr&"</a>"&ListSpaceStrs&"</td></tr>"
						else
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum*2&""" height="&RowSpace&"></td></tr>"
						end if
					Else
						if RsSysJsObj("FileType")=1 and MoreContentTF=True then
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum&""" align=Right><a class="""&LinkCSS&""" href=""" & GetOneClassLinkURL(RsTempClassObjs("ClassEName"),RsTempClassObjs("SaveFilePath"),RsTempClassObjs("FileExtName")) & """>"& MoreContentStr&"</a>"&ListSpaceStrs&"</td></tr>"
						else
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum&""" height="&RowSpace&"></td></tr>"
						end if
					End If
					JSCodeStr = JSCodeStr & "</table>');"
					WriteFile SaveFilePath,FileNameStr,JSCodeStr '写文件
					Conn.Execute("Update FS_SysJs Set AddTime='"&Now()&"' where FileName='"&FileName&"'")
					CreateSysJS = true
					RsCreateObj.Close 
					Set RsCreateObj = Nothing 
				Else
					JSCodeStr = "document.write('未查询到符合条件的新闻')"
					WriteFile SaveFilePath,FileNameStr,JSCodeStr '写文件
					Conn.Execute("Update FS_SysJs Set AddTime='"&Now()&"' where FileName='"&FileName&"'")
					CreateSysJS = "文件添加或修改成功\n\n但未找到符合条件的新闻,建议您稍后再试"
				End If
			Case "WordNews"
				if RsSysJsObj("FileType")=1 then
					RsCreateSql = "Select top "&NewsNum&" * From FS_News where ClassID in ('"&ClassIDStr&"') and DelTF=0 and HeadNewsTF=0 and PicNewsTF=0 and AuditTF=1 order by AddDate desc" '文字新闻
				else
					RsCreateSql = "Select top "&NewsNum&" * From FS_News where DelTF=0 and HeadNewsTF=0 and PicNewsTF=0 and AuditTF=1 order by AddDate desc" '文字新闻
				end if
				Set RsCreateObj = Conn.Execute(RsCreateSql)
				If Not RsCreateObj.eof then
					JSCodeStr = "document.write('<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing="""&RowSpace&"""><tr>"
					for i=1 to NewsNum
					  If RsCreateObj.eof then Exit For
					  Set TempClassObj = Conn.Execute("Select ClassEName,ClassCName,SaveFilePath from FS_NewsClass where ClassID='"&RsCreateObj("ClassID")&"'")
					  If DateTF = true then
						If ShowClassTF = true then
							If RightDate = true then
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"["&TempClassObj("ClassCName")&"]"&"<a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a></td><td><div align=Right><Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</div></td>"
							Else
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"["&TempClassObj("ClassCName")&"]"&"<a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</td>"
							End If
						Else
							If RightDate = true then
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"<a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a></td><td><div align=Right><Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</div></td>"
							Else
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"<a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</td>"
							End If
						End IF
					  Else
						If ShowClassTF = true then
							  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"["&TempClassObj("ClassCName")&"]"&"<a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;"&ListSpaceStrs&"</td>"
						Else
							  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"<a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;"&ListSpaceStrs&"</td>"
						End If
					  End If
					  RsCreateObj.MoveNext
					  if i mod Cint(RowNum) = 0 or RsCreateObj.eof then
						if RightDate = true then
						  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&RowNum*2&""" height=""1"" background=""" & AvailableDoMain & RsSysJsObj("RowBetween")&"""></td></tr><tr>"
						else
						  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&RowNum&""" height=""1"" background=""" & AvailableDoMain & RsSysJsObj("RowBetween")&"""></td></tr><tr>"
						end if
					  end if
					next 
					If RsSysJsObj("FileType")=1 then
					Set RsTempClassObjs = Conn.Execute("Select SaveFilePath,ClassEName,FileExtName from FS_NewsClass where ClassID='"&ClassID&"'")
						If RsTempClassObjs.eof then
							CreateSysJS = "刷新栏目已经不存在"
							Exit Function
						End If
					End If
					If RightDate = true then
						if RsSysJsObj("FileType")=1 and MoreContentTF=True then
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum*2&""" align=Right><a class="""&LinkCSS&""" href=""" & GetOneClassLinkURL(RsTempClassObjs("ClassEName"),RsTempClassObjs("SaveFilePath"),RsTempClassObjs("FileExtName")) &""">"& MoreContentStr&"</a>"&ListSpaceStrs&"</td></tr>"
						else
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum*2&""" height="&RowSpace&"></td></tr>"
						end if
					Else
						if RsSysJsObj("FileType")=1 and MoreContentTF=True then
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum&""" align=Right><a class="""&LinkCSS&""" href=""" & GetOneClassLinkURL(RsTempClassObjs("ClassEName"),RsTempClassObjs("SaveFilePath"),RsTempClassObjs("FileExtName")) & """>"& MoreContentStr&"</a>"&ListSpaceStrs&"</td></tr>"
						else
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum&""" height="&RowSpace&"></td></tr>"
						end if
					End If
					JSCodeStr = JSCodeStr & "</table>');"
					WriteFile SaveFilePath,FileNameStr,JSCodeStr '写文件
					Conn.Execute("Update FS_SysJs Set AddTime='"&Now()&"' where FileName='"&FileName&"'")
					CreateSysJS = true
					RsCreateObj.Close 
					Set RsCreateObj = Nothing 
				Else
					JSCodeStr = "document.write('未查询到符合条件的新闻')"
					WriteFile SaveFilePath,FileNameStr,JSCodeStr '写文件
					Conn.Execute("Update FS_SysJs Set AddTime='"&Now()&"' where FileName='"&FileName&"'")
					CreateSysJS = "文件添加或修改成功\n\n但未找到符合条件的新闻,建议您稍后再试"
				End If
			Case "TitleNews"
				if RsSysJsObj("FileType")=1 then
					RsCreateSql = "Select top "&NewsNum&" * From FS_News where ClassID in ('"&ClassIDStr&"') and DelTF=0 and HeadNewsTF=1 and AuditTF=1 order by AddDate desc" '标题新闻
				else
					RsCreateSql = "Select top "&NewsNum&" * From FS_News where DelTF=0 and HeadNewsTF=1 and AuditTF=1 order by AddDate desc" '标题新闻
				end if
				Set RsCreateObj = Conn.Execute(RsCreateSql)
				If Not RsCreateObj.eof then
					JSCodeStr = "document.write('<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing="""&RowSpace&"""><tr>"
					for i=1 to NewsNum
					  If RsCreateObj.eof then Exit For
					  Set TempClassObj = Conn.Execute("Select ClassEName,ClassCName,SaveFilePath from FS_NewsClass where ClassID='"&RsCreateObj("ClassID")&"'")
					  If DateTF = true then
						If ShowClassTF = true then
							If RightDate = true then
							  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"["&TempClassObj("ClassCName")&"]"&"<a class="""&TitleCSS&""" href="""&RsCreateObj("HeadNewsPath")&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a></td><td><div align=Right><Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</div></td>"
							Else
							  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"["&TempClassObj("ClassCName")&"]"&"<a class="""&TitleCSS&""" href="""&RsCreateObj("HeadNewsPath")&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</td>"
							End If
						Else
							If RightDate = true then
							  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"<a class="""&TitleCSS&""" href="""&RsCreateObj("HeadNewsPath")&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a></td><td><div align=Right><Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</div></td>"
							Else
							  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"<a class="""&TitleCSS&""" href="""&RsCreateObj("HeadNewsPath")&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</td>"
							End If
						End IF
					  Else
						If ShowClassTF = true then
						  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"["&TempClassObj("ClassCName")&"]"&"<a class="""&TitleCSS&""" href="""&RsCreateObj("HeadNewsPath")&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;"&ListSpaceStrs&"</td>"
						Else
						  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"<a class="""&TitleCSS&""" href="""&RsCreateObj("HeadNewsPath")&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;"&ListSpaceStrs&"</td>"
						End If
					  End If
					  RsCreateObj.MoveNext
					  if i mod Cint(RowNum) = 0 or RsCreateObj.eof then
						if RightDate = true then
						  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&RowNum*2&""" height=""1"" background=""" & AvailableDoMain & RsSysJsObj("RowBetween")&"""></td></tr><tr>"
						else
						  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&RowNum&""" height=""1"" background=""" & AvailableDoMain & RsSysJsObj("RowBetween")&"""></td></tr><tr>"
						end if
					  end if
					next 
					If RsSysJsObj("FileType")=1 then
					Set RsTempClassObjs = Conn.Execute("Select SaveFilePath,ClassEName,FileExtName from FS_NewsClass where ClassID='"&ClassID&"'")
						If RsTempClassObjs.eof then
							CreateSysJS = "刷新栏目已经不存在"
							Exit Function
						End If
					End If
					If RightDate = true then
						if RsSysJsObj("FileType")=1 and MoreContentTF=True then
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum*2&""" align=Right><a class="""&LinkCSS&""" href=""" &GetOneClassLinkURL(RsTempClassObjs("ClassEName"),RsTempClassObjs("SaveFilePath"),RsTempClassObjs("FileExtName")) &""">"& MoreContentStr&"</a>"&ListSpaceStrs&"</td></tr>"
						else
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum*2&""" height="&RowSpace&"></td></tr>"
						end if
					Else
						if RsSysJsObj("FileType")=1 and MoreContentTF=True then
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum&""" align=Right><a class="""&LinkCSS&""" href="""&GetOneClassLinkURL(RsTempClassObjs("ClassEName"),RsTempClassObjs("SaveFilePath"),RsTempClassObjs("FileExtName"))&""">"& MoreContentStr&"</a>"&ListSpaceStrs&"</td></tr>"
						else
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum&""" height="&RowSpace&"></td></tr>"
						end if
					End If
					JSCodeStr = JSCodeStr & "</table>');"
					WriteFile SaveFilePath,FileNameStr,JSCodeStr '写文件
					Conn.Execute("Update FS_SysJs Set AddTime='"&Now()&"' where FileName='"&FileName&"'")
					CreateSysJS = true
					RsCreateObj.Close 
					Set RsCreateObj = Nothing 
				Else
					JSCodeStr = "document.write('未查询到符合条件的新闻')"
					WriteFile SaveFilePath,FileNameStr,JSCodeStr '写文件
					Conn.Execute("Update FS_SysJs Set AddTime='"&Now()&"' where FileName='"&FileName&"'")
					CreateSysJS = "文件添加或修改成功\n\n但未找到符合条件的新闻,建议您稍后再试"
				End If
			Case "ProclaimNews"
				if RsSysJsObj("FileType")=1 then
					RsCreateSql = "Select top "&NewsNum&" * From FS_News where ClassID in ('"&ClassIDStr&"') and DelTF=0 and ProclaimNews=1 and AuditTF=1 order by AddDate desc" '公告新闻
				else
					RsCreateSql = "Select top "&NewsNum&" * From FS_News where DelTF=0 and ProclaimNews=1 and AuditTF=1 order by AddDate desc" '公告新闻
				end if
				Set RsCreateObj = Conn.Execute(RsCreateSql)
				If Not RsCreateObj.eof then
					JSCodeStr = "document.write('<marquee onmouseout=start() onmouseover=stop() Width="&MarWidth&" Height="&MarHeight&"  scrolldelay=80 direction="&MarDirection&" scrollamount="& CInt(MarSpeed) &"><font color=red>【公告】</font>"&BrStr
					for i=1 to NewsNum
					  If RsCreateObj.eof then Exit For
					  Set TempClassObj = Conn.Execute("Select ClassEName,ClassCName,SaveFilePath from FS_NewsClass where ClassID='"&RsCreateObj("ClassID")&"'")
					  If DateTF = true then
						If ShowClassTF = true then
							If RsCreateObj("HeadNewsTF") <> 1 then
							  JSCodeStr = JSCodeStr & NaviPic &"["&TempClassObj("ClassCName")&"]"&"<a class="""&TitleCSS&""" href="""& GetOneNewsLinkURL(RsCreateObj("NewsID")) &""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span>"&ListSpaceStrs&BrStr
							  JSCodeStr = JSCodeStr & ListSpaceStrs & Left(Replace(Replace(LoseHtml(RsCreateObj("Content")),chr(13) & chr(10),""),"&nbsp;",""),100) & "..." & ListSpaceStrs&BrStr
							Else
							  JSCodeStr = JSCodeStr & NaviPic &"["&TempClassObj("ClassCName")&"]"&"<a class="""&TitleCSS&""" href="""&RsCreateObj("HeadNewsPath")&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span>"&ListSpaceStrs&BrStr
							End If
						Else
							If RsCreateObj("HeadNewsTF") <> 1 then
							  JSCodeStr = JSCodeStr & NaviPic &"<a class="""&TitleCSS&""" href=""" & GetOneNewsLinkURL(RsCreateObj("NewsID")) &""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span>"&ListSpaceStrs&BrStr
							  JSCodeStr = JSCodeStr & ListSpaceStrs & Left(Replace(Replace(LoseHtml(RsCreateObj("Content")),chr(13) & chr(10),""),"&nbsp;",""),100) & "..." & ListSpaceStrs&BrStr
							Else
							  JSCodeStr = JSCodeStr & NaviPic &"<a class="""&TitleCSS&""" href="""&RsCreateObj("HeadNewsPath")&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span>"&ListSpaceStrs&BrStr
							End If
						End IF
					  Else
						If ShowClassTF = true then
							If RsCreateObj("HeadNewsTF") <> 1 then
							  JSCodeStr = JSCodeStr & NaviPic &"["&TempClassObj("ClassCName")&"]"&"<a class="""&TitleCSS&""" href=""" & GetOneNewsLinkURL(RsCreateObj("NewsID")) &""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;"&ListSpaceStrs&BrStr
							  JSCodeStr = JSCodeStr & ListSpaceStrs & Left(Replace(Replace(LoseHtml(RsCreateObj("Content")),chr(13) & chr(10),""),"&nbsp;",""),100) & "..." & ListSpaceStrs&BrStr
							Else
							  JSCodeStr = JSCodeStr & NaviPic &"["&TempClassObj("ClassCName")&"]"&"<a class="""&TitleCSS&""" href="""&RsCreateObj("HeadNewsPath")&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;"&ListSpaceStrs&BrStr
							End If
						Else
							If RsCreateObj("HeadNewsTF") <> 1 then
							  JSCodeStr = JSCodeStr & NaviPic &"<a class="""&TitleCSS&""" href=""" & GetOneNewsLinkURL(RsCreateObj("NewsID")) &""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;"&ListSpaceStrs&BrStr
							  JSCodeStr = JSCodeStr & ListSpaceStrs & Left(Replace(Replace(LoseHtml(RsCreateObj("Content")),chr(13) & chr(10),""),"&nbsp;",""),100) & "..." & ListSpaceStrs&BrStr
							Else
							  JSCodeStr = JSCodeStr & NaviPic &"<a class="""&TitleCSS&""" href="""&RsCreateObj("HeadNewsPath")&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;"&ListSpaceStrs&BrStr
							End If
						End If
					  End If
					  RsCreateObj.MoveNext
					next 
					If RsSysJsObj("FileType")=1 then
					Set RsTempClassObjs = Conn.Execute("Select SaveFilePath,ClassEName,FileExtName from FS_NewsClass where ClassID='"&ClassID&"'")
						If RsTempClassObjs.eof then
							CreateSysJS = "刷新栏目已经不存在"
							Exit Function
						End If
					End If
					if RsSysJsObj("FileType")=1 and MoreContentTF=True then
						JSCodeStr = JSCodeStr &"<a class="""&LinkCSS&""" href="""&GetOneClassLinkURL(RsTempClassObjs("ClassEName"),RsTempClassObjs("SaveFilePath"),RsTempClassObjs("FileExtName"))&""">"& MoreContentStr&"</a>"&ListSpaceStrs
					end if
					JSCodeStr = JSCodeStr & "</marquee>');"
					WriteFile SaveFilePath,FileNameStr,JSCodeStr '写文件
					Conn.Execute("Update FS_SysJs Set AddTime='"&Now()&"' where FileName='"&FileName&"'")
					CreateSysJS = true
					RsCreateObj.Close 
					Set RsCreateObj = Nothing 
				Else
					JSCodeStr = "document.write('未查询到符合条件的新闻')"
					WriteFile SaveFilePath,FileNameStr,JSCodeStr '写文件
					Conn.Execute("Update FS_SysJs Set AddTime='"&Now()&"' where FileName='"&FileName&"'")
					CreateSysJS = "文件添加或修改成功\n\n但未找到符合条件的新闻,建议您稍后再试"
				End If
			Case Else
				if RsSysJsObj("FileType")=1 then
					RsCreateSql = "Select top "&NewsNum&" * From FS_News where ClassID in ('"&ClassIDStr&"') and DelTF=0 and AuditTF=1 order by AddDate desc" '最新新闻
				else
					RsCreateSql = "Select top "&NewsNum&" * From FS_News where DelTF=0 and AuditTF=1 order by AddDate desc" '最新新闻
				end if
				Set RsCreateObj = Conn.Execute(RsCreateSql)
				If Not RsCreateObj.eof then
					JSCodeStr = "document.write('<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing="""&RowSpace&"""><tr>"
					for i=1 to NewsNum
					  If RsCreateObj.eof then Exit For
					  Set TempClassObj = Conn.Execute("Select ClassEName,ClassCName,SaveFilePath from FS_NewsClass where ClassID='"&RsCreateObj("ClassID")&"'")
					  If DateTF = true then
						If ShowClassTF = true then
							If RightDate = true then
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"["&TempClassObj("ClassCName")&"]"&"<a class="""&TitleCSS&""" href=""" & GetOneNewsLinkURL(RsCreateObj("NewsID")) & """" & OpenMode & ">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a></td><td><div align=Right><Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</div></td>"
							Else
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"["&TempClassObj("ClassCName")&"]"&"<a class="""&TitleCSS&""" href=""" & GetOneNewsLinkURL(RsCreateObj("NewsID")) &""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</td>"
							End If
						Else
							If RightDate = true then
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"<a class="""&TitleCSS&""" href=""" & GetOneNewsLinkURL(RsCreateObj("NewsID")) &""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a></td><td><div align=Right><Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</div></td>"
							Else
								  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"<a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;<Span class="""&DateCSS&""">"&DateFormat(RsCreateObj("AddDate"),""&DateType&"")&"</Span>"&ListSpaceStrs&"</td>"
							End If
						End IF
					  Else
						If ShowClassTF = true then
							  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"["&TempClassObj("ClassCName")&"]"&"<a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;"&ListSpaceStrs&"</td>"
						Else
							  JSCodeStr = JSCodeStr &"<td>"& NaviPic &"<a class="""&TitleCSS&""" href="""&GetOneNewsLinkURL(RsCreateObj("NewsID"))&""""&OpenMode&">"&GotTopic(LoseHtml(RsCreateObj("Title")),TitleNum)&"</a>&nbsp;"&ListSpaceStrs&"</td>"
						End If
					  End If
					  RsCreateObj.MoveNext
					  if i mod Cint(RowNum) = 0 or RsCreateObj.eof then
						if RightDate = true then
						  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&RowNum*2&""" height=""1"" background=""" & AvailableDoMain & RsSysJsObj("RowBetween")&"""></td></tr><tr>"
						else
						  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&RowNum&""" height=""1"" background=""" & AvailableDoMain & RsSysJsObj("RowBetween")&"""></td></tr><tr>"
						end if
					  end if
					next 
					If RsSysJsObj("FileType")=1 then
					Set RsTempClassObjs = Conn.Execute("Select SaveFilePath,ClassEName,FileExtName from FS_NewsClass where ClassID='"&ClassID&"'")
						If RsTempClassObjs.eof then
							CreateSysJS = "刷新栏目已经不存在"
							Exit Function
						End If
					End If
					If RightDate = true then
						if RsSysJsObj("FileType")=1 and MoreContentTF=True then
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum*2&""" align=Right><a class="""&LinkCSS&""" href=""" & GetOneNewsLinkURL(RsTempClassObjs("ClassEName"),RsTempClassObjs("SaveFilePath"),RsTempClassObjs("FileExtName")) & """>"& MoreContentStr&"</a>"&ListSpaceStrs&"</td></tr>"
						else
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum*2&""" height="&RowSpace&"></td></tr>"
						end if
					Else
						if RsSysJsObj("FileType")=1 and MoreContentTF=True then
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum&""" align=Right><a class="""&LinkCSS&""" href=""" & GetOneNewsLinkURL(RsTempClassObjs("ClassEName"),RsTempClassObjs("SaveFilePath"),RsTempClassObjs("FileExtName")) &""">"& MoreContentStr&"</a>"&ListSpaceStrs&"</td></tr>"
						else
							JSCodeStr = JSCodeStr &"<tr><td colspan="""&RowNum&""" height="&RowSpace&"></td></tr>"
						end if
					End If
					JSCodeStr = JSCodeStr & "</table>');"
					WriteFile SaveFilePath,FileNameStr,JSCodeStr '写文件
					Conn.Execute("Update FS_SysJs Set AddTime='"&Now()&"' where FileName='"&FileName&"'")
					CreateSysJS = true
					RsCreateObj.Close 
					Set RsCreateObj = Nothing 
				Else
					JSCodeStr = "document.write('未查询到符合条件的新闻')"
					WriteFile SaveFilePath,FileNameStr,JSCodeStr '写文件
					Conn.Execute("Update FS_SysJs Set AddTime='"&Now()&"' where FileName='"&FileName&"'")
					CreateSysJS = "文件添加或修改成功\n\n但未找到符合条件的新闻,建议您稍后再试"
				End If
		End Select
	Else
		CreateSysJS = "参数传递错误"
	End If
	RsSysJsObj.Close
	Set RsSysJsObj = Nothing
End Function

Function WriteFile(SaveFilePath,FileNameStr,JSCodeStr)
	Dim MyFile,CrHNJS
	Set MyFile=Server.CreateObject(G_FS_FSO)
	If MyFile.FolderExists(Server.MapPath(TempJSSysRootDir&SaveFilePath))=false then
		MyFile.CreateFolder(Server.MapPath(TempJSSysRootDir&SaveFilePath))
	End If
	If MyFile.FileExists(Server.MapPath(TempJSSysRootDir&SaveFilePath)&"/"& FileNameStr &".js") then
		MyFile.DeleteFile(Server.MapPath(TempJSSysRootDir&SaveFilePath)&"/"& FileNameStr &".js")
	End if
	Set CrHNJS=MyFile.CreateTextFile(Server.MapPath(TempJSSysRootDir & SaveFilePath) &"/"& FileNameStr &".js")
		CrHNJS.write JSCodeStr
	Set MyFile=nothing
End Function

%>