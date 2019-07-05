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
Class JSClass
	Private TempSysRootDir
	Private RsFreeJsConfig
	Private RowSpace,ListSpace,ListSpaceStr,Temp_i,TableCellSpace,TitleSpace,TitleSpaceStr,MoreContentStr
	Private AvailableDoMain
	Private Sub Class_initialize() 
		Set RsFreeJsConfig = Conn.Execute("Select DoMain from FS_Config")
		TitleSpace = 3 '新闻内容抬头空格字符个数 
		TitleSpaceStr = ""
		for Temp_i = 1 to TitleSpace
			TitleSpaceStr = TitleSpaceStr & "&nbsp;"
		next 
		AvailableDoMain = RsFreeJsConfig("DoMain")
	End Sub 
	
	Public Property Let SysRootDir(ExteriorValue)
		TempSysRootDir = ExteriorValue
	End Property
	
	Private Sub Class_Terminate()
		Set RsFreeJsConfig = Nothing
	End Sub 
	
	Private Function GetOneNewsLinkURL(NewsID)
		Dim DoMain,TempParentID,RsParentObj,RsDoMainObj,ReturnValue
		Dim CheckRootClassIndex,CheckRootClassNumber,TempClassSaveFilePath
		Dim NewsSql,RsNewsObj
		'-----------------------/l
		dim DatePathStr
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
					Set RsParentObj = Conn.Execute("Select ParentID,Domain from FS_NewsClass where ClassID='" & RsNewsObj("ParentID") & "'")
					if Not RsParentObj.Eof then
						CheckRootClassIndex = 1
						TempParentID = RsParentObj("ParentID")
						do while Not (TempParentID = "0")
							CheckRootClassIndex = CheckRootClassIndex + 1
							RsParentObj.Close
							Set RsParentObj = Nothing
							Set RsParentObj = Conn.Execute("Select ParentID,Domain from FS_NewsClass where ClassID='" & TempParentID & "'")
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
						Set RsParentObj = Nothing
					else
						Set RsParentObj = Nothing
						Set RsNewsObj = Nothing
						GetOneNewsLinkURL = ""
						Exit Function
					end if
				else
					DoMain = RsNewsObj("DoMain")
				end if
				'---------------/l
				If Application("UseDatePath")="1" Then DatePathStr=RsNewsObj("Path") else DatePathStr=""
				if (Not IsNull(DoMain)) And (DoMain <> "") then
					ReturnValue = "http://" & DoMain & "/" & RsNewsObj("ClassEName") & DatePathStr &"/" & RsNewsObj("FileName") & "." & RsNewsObj("NewsFileExtName")
				else
					if RsNewsObj("SaveFilePath") = "/" then
						TempClassSaveFilePath = RsNewsObj("SaveFilePath")
					else
						TempClassSaveFilePath = RsNewsObj("SaveFilePath") & "/"
					end if
					ReturnValue = AvailableDoMain & TempClassSaveFilePath & RsNewsObj("ClassEName") & DatePathStr & "/" & RsNewsObj("FileName") & "." & RsNewsObj("NewsFileExtName")
				end if
				'------------------/l
			end if
		end if
		Set RsNewsObj = Nothing
		GetOneNewsLinkURL = ReturnValue
	End Function
	
	Public Function WCssA(EName,CreateFileTF)
		Dim ClsJSObj,ClsJSFileObj,ClsFileSql,ClsNewsObj,TempEName,JSCodeStr,i,MyFile,CrHNJS,OpenMode
		Set ClsJSObj = Conn.Execute("Select * From FS_FreeJS where EName='"&EName&"'")
			If Not ClsJSObj.eof then
			  JSCodeStr = "document.write('<table class="""&ClsJSObj("BackCSS")&""" width=""100%"" border=""0"" cellpadding=""0"" cellspacing="""&ClsJSObj("RowSpace")&"""><tr>"
			  Set ClsJSFileObj=server.createobject(G_FS_RS)
			  ClsFileSql="select * from FS_FreeJsFile where JSName='"&EName&"' and DelFlag=0 order by ToJsTime desc"
			  ClsJSFileObj.open ClsFileSql,Conn,1,1
			  If ClsJSFileObj.eof then
				JSCodeStr = JSCodeStr & "<td>此JS内暂无新闻</td>"
			  End If
			  If ClsJSObj("OpenMode")=1 then
				  OpenMode = "target=_blank"
			  Else
				  OpenMode = "target=_self"
			  End If
			  for i=1 to ClsJSObj("NewsNum")
				  If ClsJSFileObj.eof then Exit For
				  Set ClsNewsObj = Conn.Execute("Select * From FS_News where FileName='"&ClsJSFileObj("FileName")&"'")
				  If ClsJSObj("ShowTimeTF")="1" then
					  JSCodeStr = JSCodeStr &"<td valign=middle >"&ClsJSObj("NaviPic")&"<a class="""&ClsJSObj("TitleCSS")&""" href=" & GetOneNewsLinkURL(ClsNewsObj("NewsID")) &" "&OpenMode&">"&GotTopic(ClsNewsObj("Title"),ClsJSObj("NewsTitleNum"))&"</a></td><td><Span class="""&ClsJSObj("DateCSS")&""">"&DateFormat(ClsNewsObj("AddDate"),""&ClsJSObj("DateType")&"")&"</Span></td>"
				  Else
					  JSCodeStr = JSCodeStr &"<td valign=middle>"&ClsJSObj("NaviPic")&"<a class="""&ClsJSObj("TitleCSS")&""" href=" & GetOneNewsLinkURL(ClsNewsObj("NewsID")) &" "&OpenMode&">"&GotTopic(ClsNewsObj("Title"),ClsJSObj("NewsTitleNum"))&"</a></td>"
				  End If
				  ClsNewsObj.Close
				  Set ClsNewsObj = Nothing
				  ClsJSFileObj.MoveNext
				  if i mod Cint(ClsJSObj("RowNum")) = 0 or ClsJSFileObj.eof then
					if ClsJSObj("ShowTimeTF")=1 then
						  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))*2&""" height="""&ClsJSObj("RowSpace")&""" background="""& AvailableDoMain & ClsJSObj("RowBettween")&"""></td></tr><tr>"
					else
						  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))&""" height="""&ClsJSObj("RowSpace")&""" background="""& AvailableDoMain & ClsJSObj("RowBettween")&"""></td></tr><tr>"
					end if
				  end if
			  next 
			  ClsJSFileObj.Close 
			  Set ClsJSFileObj = Nothing 
			  JSCodeStr = JSCodeStr & "</tr></table>');"
			  JSCodeStr = Replace(JSCodeStr,"<tr></tr>","")
			  JSCodeStr = Replace(Replace(JSCodeStr,Chr(13),""),Chr(10),"")
			  Set MyFile=Server.CreateObject(G_FS_FSO)
			  If MyFile.FileExists(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js") then
				 MyFile.DeleteFile(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js")
			  End if
			  Set CrHNJS=MyFile.CreateTextFile(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js")
				  CrHNJS.write JSCodeStr
			  Set MyFile=nothing
			  '---------
			  ClsJSObj.Close
			  Set ClsJSObj = Nothing
			Else
				WCssA = "生成JS文件时未找到参数"
			End If
	End Function 

	Public Function WCssB(EName,CreateFileTF)
		Dim ClsJSObj,ClsJSFileObj,ClsFileSql,ClsNewsObj,TempEName,JSCodeStr,i,MyFile,CrHNJS,OpenMode
		Dim NewsLinkStr
		Set ClsJSObj = Conn.Execute("Select * From FS_FreeJS where EName='"&EName&"'")
			If Not ClsJSObj.eof then
			  JSCodeStr = "document.write('<table class="""&ClsJSObj("BackCSS")&""" width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr>"
			  Set ClsJSFileObj=server.createobject(G_FS_RS)
			  ClsFileSql="select * from FS_FreeJsFile where JSName='"&EName&"' and DelFlag=0 order by ToJsTime desc"
			  ClsJSFileObj.open ClsFileSql,Conn,1,1
			  If ClsJSFileObj.eof then 
				JSCodeStr = JSCodeStr & "<td>此JS内暂无新闻</td>"
			  end if
			  If ClsJSObj("OpenMode")=1 then
				  OpenMode = "target=_blank"
			  Else
				  OpenMode = "target=_self"
			  End If
				ListSpaceStr = ""
				for Temp_i = 1 to Cint(ClsJSObj("RowSpace"))
					ListSpaceStr = ListSpaceStr & "&nbsp;"
				next 
			  for i=1 to ClsJSObj("NewsNum")
				  If ClsJSFileObj.eof then Exit For
				  Set ClsNewsObj = Conn.Execute("Select * From FS_News where FileName='"&ClsJSFileObj("FileName")&"'")
				  NewsLinkStr = GetOneNewsLinkURL(ClsNewsObj("NewsID"))
				  If ClsJSObj("ShowTimeTF")=1 then
					  JSCodeStr = JSCodeStr &"<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td>"& ClsJSObj("NaviPic") &"<a class="""&ClsJSObj("TitleCSS")&""" href=" & NewsLinkStr &" "&OpenMode&">"&GotTopic(ClsNewsObj("Title"),ClsJSObj("NewsTitleNum"))&"</a></td><td><Span class="""&ClsJSObj("DateCSS")&""">"&DateFormat(ClsNewsObj("AddDate"),""&ClsJSObj("DateType")&"")&"</Span></td><td rowspan=2>"&ListSpaceStr&"</td></tr>"
				  Else
					  JSCodeStr = JSCodeStr &"<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td>"& ClsJSObj("NaviPic") &"<a class="""&ClsJSObj("TitleCSS")&""" href=" & NewsLinkStr &" "&OpenMode&">"&GotTopic(ClsNewsObj("Title"),ClsJSObj("NewsTitleNum"))&"</a></td><td rowspan=2>"&ListSpaceStr&"</td></tr>"
				  End If
				  If ClsJSObj("ShowTimeTF")=1 then
					If ClsJSObj("MoreContent")=1 then
					  JSCodeStr = JSCodeStr & "<tr><td colspan=2><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......<br><div align=""right""><a class="""&ClsJSObj("LinkCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&ClsJSObj("LinkWord")&"</a></div></td></tr></table></td>"
					Else
					  JSCodeStr = JSCodeStr & "<tr><td colspan=2><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......</td></tr></table></td>"
					End If
				  Else
					If ClsJSObj("MoreContent")=1 then
					  JSCodeStr = JSCodeStr & "<tr><td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......<br><div align=""right""><a class="""&ClsJSObj("LinkCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&ClsJSObj("LinkWord")&"</a></div></td></tr></table></td>"
					Else
					  JSCodeStr = JSCodeStr & "<tr><td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......</td></tr></table></td>"
					End If
				  End If
				  ClsNewsObj.Close
				  Set ClsNewsObj = Nothing
				  ClsJSFileObj.MoveNext
				  if i mod Cint(ClsJSObj("RowNum")) = 0 or ClsJSFileObj.eof then
					  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))&""" height="""&ClsJSObj("RowSpace")&""" background="""&ClsJSObj("RowBettween")&"""></td></tr><tr>"
				  end if
			  next 
			  ClsJSFileObj.Close 
			  Set ClsJSFileObj = Nothing 
			  JSCodeStr = JSCodeStr & "</tr></table>');"
			  JSCodeStr = Replace(JSCodeStr,"<tr></tr>","")
			  JSCodeStr = Replace(Replace(JSCodeStr,Chr(13),""),Chr(10),"")
			  if CreateFileTF = True then
				  Set MyFile=Server.CreateObject(G_FS_FSO)
				  If MyFile.FileExists(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js") then
					 MyFile.DeleteFile(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js")
				  End if
				  Set CrHNJS=MyFile.CreateTextFile(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js")
					  CrHNJS.write JSCodeStr
				  Set MyFile=nothing
				  ClsJSObj.Close
				  Set ClsJSObj = Nothing
			  else
				  WCssB = JSCodeStr
			  end if
			Else
				WCssB = "生成JS文件时未找到参数"
			End If
	End Function 

	Public Function WCssC(EName,CreateFileTF)
		Dim ClsJSObj,ClsJSFileObj,ClsFileSql,ClsNewsObj,TempEName,JSCodeStr,i,MyFile,CrHNJS,OpenMode
		Dim NewsLinkStr
		Set ClsJSObj = Conn.Execute("Select * From FS_FreeJS where EName='"&EName&"'")
			If Not ClsJSObj.eof then
			  JSCodeStr = "document.write('<table class="""&ClsJSObj("BackCSS")&""" width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr>"
			  Set ClsJSFileObj=server.createobject(G_FS_RS)
			  ClsFileSql="select * from FS_FreeJsFile where JSName='"&EName&"' and DelFlag=0 order by ToJsTime desc"
			  ClsJSFileObj.open ClsFileSql,Conn,1,1
			  If ClsJSFileObj.eof then 
				JSCodeStr = JSCodeStr & "<td>此JS内暂无新闻</td></tr>"
			  end if
			  If ClsJSObj("OpenMode")=1 then
				  OpenMode = "target=_blank"
			  Else
				  OpenMode = "target=_self"
			  End If
				ListSpaceStr = ""
				for Temp_i = 1 to Cint(ClsJSObj("RowSpace"))
					ListSpaceStr = ListSpaceStr & "&nbsp;"
				next 
			  for i=1 to ClsJSObj("NewsNum")
				  If ClsJSFileObj.eof then Exit For
				  Set ClsNewsObj = Conn.Execute("Select * From FS_News where FileName='"&ClsJSFileObj("FileName")&"'")
				  NewsLinkStr = GetOneNewsLinkURL(ClsNewsObj("NewsID"))
				  If ClsJSObj("ShowTimeTF")="1" then
					  JSCodeStr = JSCodeStr &"<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td valign=""top""><div align=""center"">"& ClsJSObj("NaviPic") &"<br><Span class="""&ClsJSObj("TitleCSS")&""">"&ListTitle(GotTopic(ClsNewsObj("Title"),ClsJSObj("NewsTitleNum")),ClsJSObj("NewsTitleNum"))&"</Span><br><Span class="""&ClsJSObj("DateCSS")&""">"&ListTitle(DateFormat(ClsNewsObj("AddDate"),""&ClsJSObj("DateType")&""),50)&"</Span></div></td>"
				  Else
					  JSCodeStr = JSCodeStr &"<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td valign=""top""><div align=""center"">"& ClsJSObj("NaviPic") &"<br><Span class="""&ClsJSObj("TitleCSS")&""">"&ListTitle(GotTopic(ClsNewsObj("Title"),ClsJSObj("NewsTitleNum")),ClsJSObj("NewsTitleNum"))&"</Span></div></td>"
				  End If
				  If ClsJSObj("MoreContent")=1 then
					  JSCodeStr = JSCodeStr & "<td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......<br><div align=""right""><a class="""&ClsJSObj("LinkCSS")&""" href=" & NewsLinkStr&"."&ClsNewsObj("FileExtName")&" "&OpenMode&">"&ClsJSObj("LinkWord")&"</a></div></td><td>"&ListSpaceStr&"</td></tr></table></td>"
				  Else
					  JSCodeStr = JSCodeStr & "<td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......</td><td>"&ListSpaceStr&"</td></tr></table></td>"
				  End If
				  ClsNewsObj.Close
				  Set ClsNewsObj = Nothing
				  ClsJSFileObj.MoveNext
				  if i mod Cint(ClsJSObj("RowNum")) = 0 or ClsJSFileObj.eof then
					  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))&""" height="""&ClsJSObj("RowSpace")&""" background="""&ClsJSObj("RowBettween")&"""></td></tr><tr>"
				  end if
			  next 
			  ClsJSFileObj.Close 
			  Set ClsJSFileObj = Nothing 
			  JSCodeStr = JSCodeStr & "</tr></table>');"
			  JSCodeStr = Replace(JSCodeStr,"<tr></tr>","")
			  JSCodeStr = Replace(Replace(JSCodeStr,Chr(13),""),Chr(10),"")
			  if CreateFileTF = True then
				  Set MyFile=Server.CreateObject(G_FS_FSO)
				  If MyFile.FileExists(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js") then
					 MyFile.DeleteFile(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js")
				  End if
				  Set CrHNJS=MyFile.CreateTextFile(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js")
					  CrHNJS.write JSCodeStr
				  Set MyFile=nothing
				  ClsJSObj.Close
				  Set ClsJSObj = Nothing
				else
					WCssC = JSCodeStr
				end if
			Else
				WCssC = "生成JS文件时未找到参数"
			End If
	End Function 

	Public Function WCssD(EName,CreateFileTF)
		Dim ClsJSObj,ClsJSFileObj,ClsFileSql,ClsNewsObj,TempEName,JSCodeStr,i,MyFile,CrHNJS,OpenMode
		Dim NewsLinkStr
		Set ClsJSObj = Conn.Execute("Select * From FS_FreeJS where EName='"&EName&"'")
			If Not ClsJSObj.eof then
			  JSCodeStr = "document.write('<table class="""&ClsJSObj("BackCSS")&""" width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr>"
			  Set ClsJSFileObj=server.createobject(G_FS_RS)
			  ClsFileSql="select * from FS_FreeJsFile where JSName='"&EName&"' and DelFlag=0 order by ToJsTime desc"
			  ClsJSFileObj.open ClsFileSql,Conn,1,1
			  If ClsJSFileObj.eof then 
				JSCodeStr = JSCodeStr & "<td>此JS内暂无新闻</td></tr>"
			  end if
			  If ClsJSObj("OpenMode")=1 then
				  OpenMode = "target=_blank"
			  Else
				  OpenMode = "target=_self"
			  End If
				ListSpaceStr = ""
				for Temp_i = 1 to Cint(ClsJSObj("RowSpace"))
					ListSpaceStr = ListSpaceStr & "&nbsp;"
				next 
			  for i=1 to ClsJSObj("NewsNum")
				  If ClsJSFileObj.eof then Exit For
				  Set ClsNewsObj = Conn.Execute("Select * From FS_News where FileName='"&ClsJSFileObj("FileName")&"'")
				  NewsLinkStr = GetOneNewsLinkURL(ClsNewsObj("NewsID"))
				  If ClsJSObj("MoreContent")=1 then
					  JSCodeStr = JSCodeStr & "<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td valign=""top""><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......<br><div align=""right""><a class="""&ClsJSObj("LinkCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&ClsJSObj("LinkWord")&"</a></div></td>"
				  Else
					  JSCodeStr = JSCodeStr & "<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td valign=""top""><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......</td>"
				  End If
				  If ClsJSObj("ShowTimeTF")="1" then
					  JSCodeStr = JSCodeStr &"<td valign=""top""><div align=""center"">"& ClsJSObj("NaviPic") &"<br><Span class="""&ClsJSObj("TitleCSS")&""">"&ListTitle(GotTopic(ClsNewsObj("Title"),ClsJSObj("NewsTitleNum")),ClsJSObj("NewsTitleNum"))&"</Span><br><Span class="""&ClsJSObj("DateCSS")&""">"&ListTitle(DateFormat(ClsNewsObj("AddDate"),""&ClsJSObj("DateType")&""),50)&"</Span></div></td><td>"&ListSpaceStr&"</td></tr></table></td>"
				  Else
					  JSCodeStr = JSCodeStr &"<td valign=""top""><div align=""center"">"& ClsJSObj("NaviPic") &"<br><Span class="""&ClsJSObj("TitleCSS")&""">"&ListTitle(GotTopic(ClsNewsObj("Title"),ClsJSObj("NewsTitleNum")),ClsJSObj("NewsTitleNum"))&"</Span></div></td><td>"&ListSpaceStr&"</td></tr></table></td>"
				  End If
				  ClsNewsObj.Close
				  Set ClsNewsObj = Nothing
				  ClsJSFileObj.MoveNext
				  if i mod Cint(ClsJSObj("RowNum")) = 0 or ClsJSFileObj.eof then
					  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))&""" height="""&ClsJSObj("RowSpace")&""" background="""&ClsJSObj("RowBettween")&"""></td></tr><tr>"
				  end if
			  next 
			  ClsJSFileObj.Close 
			  Set ClsJSFileObj = Nothing 
			  JSCodeStr = JSCodeStr & "</tr></table>');"
			  JSCodeStr = Replace(JSCodeStr,"<tr></tr>","")
			  JSCodeStr = Replace(Replace(JSCodeStr,Chr(13),""),Chr(10),"")
			  if CreateFileTF = True then
				  Set MyFile=Server.CreateObject(G_FS_FSO)
				  If MyFile.FileExists(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js") then
					 MyFile.DeleteFile(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js")
				  End if
				  Set CrHNJS=MyFile.CreateTextFile(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js")
					  CrHNJS.write JSCodeStr
				  Set MyFile=nothing
				  ClsJSObj.Close
				  Set ClsJSObj = Nothing
				else
					WCssD = JSCodeStr
				end if
			Else
				WCssD = "生成JS文件时未找到参数"
			End If
	End Function

	Public Function WCssE(EName,CreateFileTF)
		Dim ClsJSObj,ClsJSFileObj,ClsFileSql,ClsNewsObj,TempEName,JSCodeStr,i,MyFile,CrHNJS,OpenMode
		Dim NewsLinkStr
		Set ClsJSObj = Conn.Execute("Select * From FS_FreeJS where EName='"&EName&"'")
			If Not ClsJSObj.eof then
			  JSCodeStr = "document.write('<table class="""&ClsJSObj("BackCSS")&""" width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr>"
			  Set ClsJSFileObj=server.createobject(G_FS_RS)
			  ClsFileSql="select * from FS_FreeJsFile where JSName='"&EName&"' and DelFlag=0 order by ToJsTime desc"
			  ClsJSFileObj.open ClsFileSql,Conn,1,1
			  If ClsJSFileObj.eof then 
				JSCodeStr = JSCodeStr & "<td>此JS内暂无新闻</td></tr>"
			  end if
			  If ClsJSObj("OpenMode")=1 then
				  OpenMode = "target=_blank"
			  Else
				  OpenMode = "target=_self"
			  End If
				ListSpaceStr = ""
				for Temp_i = 1 to Cint(ClsJSObj("RowSpace"))
					ListSpaceStr = ListSpaceStr & "&nbsp;"
				next 
			  for i=1 to ClsJSObj("NewsNum")
				  If ClsJSFileObj.eof then Exit For
				  Set ClsNewsObj = Conn.Execute("Select * From FS_News where FileName='"&ClsJSFileObj("FileName")&"'")
				  NewsLinkStr = GetOneNewsLinkURL(ClsNewsObj("NewsID"))
				  If ClsJSObj("MoreContent")=1 then
					  JSCodeStr = JSCodeStr & "<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......<br><div align=""right""><a class="""&ClsJSObj("LinkCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&ClsJSObj("LinkWord")&"</a></div></td><td rowspan=2>"&ListSpaceStr&"</td></tr>"
				  Else
					  JSCodeStr = JSCodeStr & "<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......</td><td rowspan=2>"&ListSpaceStr&"</td></tr>"
				  End If
				  If ClsJSObj("ShowTimeTF")=1 then
					  JSCodeStr = JSCodeStr &"<tr><td>"& ClsJSObj("NaviPic") &"<a class="""&ClsJSObj("TitleCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&GotTopic(ClsNewsObj("Title"),ClsJSObj("NewsTitleNum"))&"</a>&nbsp;<Span class="""&ClsJSObj("DateCSS")&""">"&DateFormat(ClsNewsObj("AddDate"),""&ClsJSObj("DateType")&"")&"</Span></td></tr></table></td>"
				  Else
					  JSCodeStr = JSCodeStr &"<tr><td>"& ClsJSObj("NaviPic") &"<a class="""&ClsJSObj("TitleCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&GotTopic(ClsNewsObj("Title"),ClsJSObj("NewsTitleNum"))&"</a></td></tr></table></td>"
				  End If
				  ClsNewsObj.Close
				  Set ClsNewsObj = Nothing
				  ClsJSFileObj.MoveNext
				  if i mod Cint(ClsJSObj("RowNum")) = 0 or ClsJSFileObj.eof then
					  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))&""" height="""&ClsJSObj("RowSpace")&""" background="""&ClsJSObj("RowBettween")&"""></td></tr><tr>"
				  end if
			  next 
			  ClsJSFileObj.Close 
			  Set ClsJSFileObj = Nothing 
			  JSCodeStr = JSCodeStr & "</tr></table>');"
			  JSCodeStr = Replace(JSCodeStr,"<tr></tr>","")
			  JSCodeStr = Replace(Replace(JSCodeStr,Chr(13),""),Chr(10),"")
			  if CreateFileTF = True then
				  Set MyFile=Server.CreateObject(G_FS_FSO)
				  If MyFile.FileExists(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js") then
					 MyFile.DeleteFile(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js")
				  End if
				  Set CrHNJS=MyFile.CreateTextFile(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js")
					  CrHNJS.write JSCodeStr
				  Set MyFile=nothing
				  ClsJSObj.Close
				  Set ClsJSObj = Nothing
			  else
				  WCssE = JSCodeStr
			  end if
			Else
				WCssE = "生成JS文件时未找到参数"
			End If
	End Function

	Public Function PCssA(EName,CreateFileTF)
		Dim ClsJSObj,ClsJSFileObj,ClsFileSql,ClsNewsObj,TempEName,JSCodeStr,i,MyFile,CrHNJS,OpenMode
		Dim NewsLinkStr
		Set ClsJSObj = Conn.Execute("Select * From FS_FreeJS where EName='"&EName&"'")
			If Not ClsJSObj.eof then
			  JSCodeStr = "document.write('<table class="""&ClsJSObj("BackCSS")&""" width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr>"
			  Set ClsJSFileObj=server.createobject(G_FS_RS)
			  ClsFileSql="select * from FS_FreeJsFile where JSName='"&EName&"' and DelFlag=0 order by ToJsTime desc"
			  ClsJSFileObj.open ClsFileSql,Conn,1,1
			  If ClsJSFileObj.eof then 
				JSCodeStr = JSCodeStr & "<td>此JS内暂无新闻</td></tr>"
			  end if
			  If ClsJSObj("OpenMode")=1 then
				  OpenMode = "target=_blank"
			  Else
				  OpenMode = "target=_self"
			  End If
			  for i=1 to ClsJSObj("NewsNum")
				  If ClsJSFileObj.eof then Exit For
				  Set ClsNewsObj = Conn.Execute("Select * From FS_News where FileName='"&ClsJSFileObj("FileName")&"'")
				  NewsLinkStr = GetOneNewsLinkURL(ClsNewsObj("NewsID"))
				  JSCodeStr = JSCodeStr & "<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td align=""center""><a href="&NewsLinkStr&" "&OpenMode&"><img src="& AvailableDoMain & ClsJSFileObj("PicPath") &" height="&ClsJSObj("PicHeight")&" width="&ClsJSObj("PicWidth")&" border=""0""></a></td></tr>"
				  If ClsJSObj("ShowTimeTF")="1" then
					  JSCodeStr = JSCodeStr & "<tr><td align=""center"">"& ClsJSObj("NaviPic") &"<a class="""&ClsJSObj("TitleCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&GotTopic(ClsNewsObj("Title"),ClsJSObj("NewsTitleNum"))&"</a>&nbsp;<Span class="""&ClsJSObj("DateCSS")&""">"&DateFormat(ClsNewsObj("AddDate"),""&ClsJSObj("DateType")&"")&"</Span></td></tr></table></td>"
				  Else
					  JSCodeStr = JSCodeStr & "<tr><td align=""center"">"& ClsJSObj("NaviPic") &"<a class="""&ClsJSObj("TitleCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&GotTopic(ClsNewsObj("Title"),ClsJSObj("NewsTitleNum"))&"</a></td></tr></table></td>"
				  End If
				  ClsNewsObj.Close
				  Set ClsNewsObj = Nothing
				  ClsJSFileObj.MoveNext
				  if i mod Cint(ClsJSObj("RowNum")) = 0 or ClsJSFileObj.eof then
					  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))&""" height="""&ClsJSObj("RowSpace")&""" background="""&ClsJSObj("RowBettween")&"""></td></tr><tr>"
				  end if
			  next 
			  ClsJSFileObj.Close 
			  Set ClsJSFileObj = Nothing 
			  JSCodeStr = JSCodeStr & "</tr></table>');"
			  JSCodeStr = Replace(JSCodeStr,"<tr></tr>","")
			  JSCodeStr = Replace(Replace(JSCodeStr,Chr(13),""),Chr(10),"")
			  if CreateFileTF = True then
				  Set MyFile=Server.CreateObject(G_FS_FSO)
				  If MyFile.FileExists(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js") then
					 MyFile.DeleteFile(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js")
				  End if
				  Set CrHNJS=MyFile.CreateTextFile(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js")
					  CrHNJS.write JSCodeStr
				  Set MyFile=nothing
				  ClsJSObj.Close
				  Set ClsJSObj = Nothing
			  else
				  PCssA = JSCodeStr
			  end if
			Else
				PCssA = "生成JS文件时未找到参数"
			End If
	End Function

	Public Function PCssB(EName,CreateFileTF)
		Dim ClsJSObj,ClsJSFileObj,ClsFileSql,ClsNewsObj,TempEName,JSCodeStr,i,MyFile,CrHNJS,OpenMode
		Dim NewsLinkStr
		Set ClsJSObj = Conn.Execute("Select * From FS_FreeJS where EName='"&EName&"'")
			If Not ClsJSObj.eof then
			  JSCodeStr = "document.write('<table class="""&ClsJSObj("BackCSS")&""" width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr>"
			  Set ClsJSFileObj=server.createobject(G_FS_RS)
			  ClsFileSql="select * from FS_FreeJsFile where JSName='"&EName&"' and DelFlag=0 order by ToJsTime desc"
			  ClsJSFileObj.open ClsFileSql,Conn,1,1
			  If ClsJSFileObj.eof then 
				JSCodeStr = JSCodeStr & "<td>此JS内暂无新闻</td></tr>"
			  end if
			  If ClsJSObj("OpenMode")=1 then
				  OpenMode = "target=_blank"
			  Else
				  OpenMode = "target=_self"
			  End If
				ListSpaceStr = ""
				for Temp_i = 1 to Cint(ClsJSObj("RowSpace")\2)
					ListSpaceStr = ListSpaceStr & "&nbsp;"
				next 
			  for i=1 to ClsJSObj("NewsNum")
				  If ClsJSFileObj.eof then Exit For
				  Set ClsNewsObj = Conn.Execute("Select * From FS_News where FileName='"&ClsJSFileObj("FileName")&"'")
				  NewsLinkStr = GetOneNewsLinkURL(ClsNewsObj("NewsID"))
				  JSCodeStr = JSCodeStr & "<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td valign=""top"" align=""center"" rowspan=""2""><a href="&NewsLinkStr&" "&OpenMode&"><img src="& AvailableDoMain & ClsJSFileObj("PicPath") &" height="&ClsJSObj("PicHeight")&" width="&ClsJSObj("PicWidth")&" border=""0""></a></td><td rowspan=""2"">"&ListSpaceStr&"</td>"
				  If ClsJSObj("ShowTimeTF")="1" then
					  JSCodeStr = JSCodeStr & "<td align=""left"">"& ClsJSObj("NaviPic") &"<a class="""&ClsJSObj("TitleCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&GotTopic(ClsNewsObj("Title"),ClsJSObj("NewsTitleNum"))&"</a>&nbsp;<Span class="""&ClsJSObj("DateCSS")&""">"&DateFormat(ClsNewsObj("AddDate"),""&ClsJSObj("DateType")&"")&"</Span></td><td rowspan=""2"">"&ListSpaceStr&"</td></tr>"
				  Else
				  response.write ClsNewsObj("Title")&"---"&ClsJSObj("NewsTitleNum")
					  JSCodeStr = JSCodeStr & "<td align=""left"">"& ClsJSObj("NaviPic") &"<a class="""&ClsJSObj("TitleCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&GotTopic(ClsNewsObj("Title"),ClsJSObj("NewsTitleNum"))&"</a></td><td rowspan=""2"">"&ListSpaceStr&"</td></tr>"
				  End If
				  If ClsJSObj("MoreContent")=1 then
					  JSCodeStr = JSCodeStr & "<tr><td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......<br><div align=""right""><a class="""&ClsJSObj("LinkCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&ClsJSObj("LinkWord")&"</a></div></td></tr></table></td>"
				  Else
					  JSCodeStr = JSCodeStr & "<tr><td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......</td></tr></table></td>"
				  End If
				  ClsNewsObj.Close
				  Set ClsNewsObj = Nothing
				  ClsJSFileObj.MoveNext
				  if i mod Cint(ClsJSObj("RowNum")) = 0 or ClsJSFileObj.eof then
					  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))&""" height="""&ClsJSObj("RowSpace")&""" background="""&ClsJSObj("RowBettween")&"""></td></tr><tr>"
				  end if
			  next 
			  ClsJSFileObj.Close 
			  Set ClsJSFileObj = Nothing 
			  JSCodeStr = JSCodeStr & "</tr></table>');"
			  JSCodeStr = Replace(Replace(JSCodeStr,"<tr></tr>",""),"&nbsp;&nbsp;&nbsp;&nbsp;","")
			  JSCodeStr = Replace(Replace(JSCodeStr,Chr(13),""),Chr(10),"")
			  if CreateFileTF = True then
				  Set MyFile=Server.CreateObject(G_FS_FSO)
				  If MyFile.FileExists(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js") then
					 MyFile.DeleteFile(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js")
				  End if
				  Set CrHNJS=MyFile.CreateTextFile(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js")
					  CrHNJS.write JSCodeStr
				  Set MyFile=nothing
				  ClsJSObj.Close
				  Set ClsJSObj = Nothing
			   else
				  PCssB = JSCodeStr
			   end if
			Else
				PCssB = "生成JS文件时未找到参数"
			End If
	End Function

	Public Function PCssC(EName,CreateFileTF)
		Dim ClsJSObj,ClsJSFileObj,ClsFileSql,ClsNewsObj,TempEName,JSCodeStr,i,MyFile,CrHNJS,OpenMode
		Dim NewsLinkStr
		Set ClsJSObj = Conn.Execute("Select * From FS_FreeJS where EName='"&EName&"'")
			If Not ClsJSObj.eof then
			  JSCodeStr = "document.write('<table class="""&ClsJSObj("BackCSS")&""" width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr>"
			  Set ClsJSFileObj=server.createobject(G_FS_RS)
			  ClsFileSql="select * from FS_FreeJsFile where JSName='"&EName&"' and DelFlag=0 order by ToJsTime desc"
			  ClsJSFileObj.open ClsFileSql,Conn,1,1
			  If ClsJSFileObj.eof then 
				JSCodeStr = JSCodeStr & "<td>此JS内暂无新闻</td></tr>"
			  end if
			  If ClsJSObj("OpenMode")=1 then
				  OpenMode = "target=_blank"
			  Else
				  OpenMode = "target=_self"
			  End If
			  for i=1 to ClsJSObj("NewsNum")
				  If ClsJSFileObj.eof then Exit For
				  Set ClsNewsObj = Conn.Execute("Select * From FS_News where FileName='"&ClsJSFileObj("FileName")&"'")
				  NewsLinkStr = GetOneNewsLinkURL(ClsNewsObj("NewsID"))
				  JSCodeStr = JSCodeStr & "<td align=""center""><a href="&NewsLinkStr&" "&OpenMode&"><img src="& AvailableDoMain & ClsJSFileObj("PicPath") &" height="&ClsJSObj("PicHeight")&" width="&ClsJSObj("PicWidth")&" border=""0""></a></td>"
				  ClsNewsObj.Close
				  Set ClsNewsObj = Nothing
				  ClsJSFileObj.MoveNext
				  if i mod Cint(ClsJSObj("RowNum")) = 0 or ClsJSFileObj.eof then
					  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))&""" height="""&ClsJSObj("RowSpace")&""" background="""&ClsJSObj("RowBettween")&"""></td></tr><tr>"
				  end if
			  next 
			  ClsJSFileObj.Close 
			  Set ClsJSFileObj = Nothing 
			  JSCodeStr = JSCodeStr & "</tr></table>');"
			  JSCodeStr = Replace(JSCodeStr,"<tr></tr>","")
			  JSCodeStr = Replace(Replace(JSCodeStr,Chr(13),""),Chr(10),"")
			  if CreateFileTF = True then
				  Set MyFile=Server.CreateObject(G_FS_FSO)
				  If MyFile.FileExists(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js") then
					 MyFile.DeleteFile(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js")
				  End if
				  Set CrHNJS=MyFile.CreateTextFile(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js")
					  CrHNJS.write JSCodeStr
				  Set MyFile=nothing
				  ClsJSObj.Close
				  Set ClsJSObj = Nothing
			   else
				  PCssC = JSCodeStr
			   end if
			Else
				PCssC = "生成JS文件时未找到参数"
			End If
	End Function

	Public Function PCssD(EName,CreateFileTF)
		Dim ClsJSObj,ClsJSFileObj,ClsFileSql,ClsNewsObj,TempEName,JSCodeStr,i,MyFile,CrHNJS,OpenMode
		Dim NewsLinkStr
		Set ClsJSObj = Conn.Execute("Select * From FS_FreeJS where EName='"&EName&"'")
			If Not ClsJSObj.eof then
			  JSCodeStr = "document.write('<table class="""&ClsJSObj("BackCSS")&""" width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr>"
			  Set ClsJSFileObj=server.createobject(G_FS_RS)
			  ClsFileSql="select * from FS_FreeJsFile where JSName='"&EName&"' and DelFlag=0 order by ToJsTime desc"
			  ClsJSFileObj.open ClsFileSql,Conn,1,1
			  If ClsJSFileObj.eof then 
				JSCodeStr = JSCodeStr & "<td>此JS内暂无新闻</td></tr>"
			  end if
			  If ClsJSObj("OpenMode")=1 then
				  OpenMode = "target=_blank"
			  Else
				  OpenMode = "target=_self"
			  End If
				ListSpaceStr = ""
				for Temp_i = 1 to Cint(ClsJSObj("RowSpace"))
					ListSpaceStr = ListSpaceStr & "&nbsp;"
				next 
			  for i=1 to ClsJSObj("NewsNum")
				  If ClsJSFileObj.eof then Exit For
				  Set ClsNewsObj = Conn.Execute("Select * From FS_News where FileName='"&ClsJSFileObj("FileName")&"'")
				  NewsLinkStr = GetOneNewsLinkURL(ClsNewsObj("NewsID"))
				  JSCodeStr = JSCodeStr & "<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td valign=""top"" colspan=""2"" align=""center""><a href="&NewsLinkStr&" "&OpenMode&"><img src="& AvailableDoMain & ClsJSFileObj("PicPath") &" height="&ClsJSObj("PicHeight")&" width="&ClsJSObj("PicWidth")&" border=""0""></a></td><td rowspan=""2"">"&ListSpaceStr&"</td></tr>"
				  If ClsJSObj("ShowTimeTF")="1" then
					  JSCodeStr = JSCodeStr &"<tr><td valign=""top""><div align=""center"">"& ClsJSObj("NaviPic") &"<br><Span class="""&ClsJSObj("TitleCSS")&""">"&ListTitle(GotTopic(ClsNewsObj("Title"),ClsJSObj("NewsTitleNum")),ClsJSObj("NewsTitleNum"))&"</Span><br><Span class="""&ClsJSObj("DateCSS")&""">"&ListTitle(DateFormat(ClsNewsObj("AddDate"),""&ClsJSObj("DateType")&""),50)&"</Span></div></td>"
				  Else
					  JSCodeStr = JSCodeStr &"<tr><td valign=""top""><div align=""center"">"& ClsJSObj("NaviPic") &"<br><Span class="""&ClsJSObj("TitleCSS")&""">"&ListTitle(GotTopic(ClsNewsObj("Title"),ClsJSObj("NewsTitleNum")),ClsJSObj("NewsTitleNum"))&"</Span></div></td>"
				  End If
				  If ClsJSObj("MoreContent")=1 then
					  JSCodeStr = JSCodeStr & "<td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......<br><div align=""right""><a class="""&ClsJSObj("LinkCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&ClsJSObj("LinkWord")&"</a></div></td></tr></table></td>"
				  Else
					  JSCodeStr = JSCodeStr & "<td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......</td></tr></table></td>"
				  End If
				  ClsNewsObj.Close
				  Set ClsNewsObj = Nothing
				  ClsJSFileObj.MoveNext
				  if i mod Cint(ClsJSObj("RowNum")) = 0 or ClsJSFileObj.eof then
					  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))&""" height="""&ClsJSObj("RowSpace")&""" background="""&ClsJSObj("RowBettween")&"""></td></tr><tr>"
				  end if
			  next 
			  ClsJSFileObj.Close 
			  Set ClsJSFileObj = Nothing 
			  JSCodeStr = JSCodeStr & "</tr></table>');"
			  JSCodeStr = Replace(JSCodeStr,"<tr></tr>","")
			  JSCodeStr = Replace(Replace(JSCodeStr,Chr(13),""),Chr(10),"")
			  if CreateFileTF = True then
				  Set MyFile=Server.CreateObject(G_FS_FSO)
				  If MyFile.FileExists(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js") then
					 MyFile.DeleteFile(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js")
				  End if
				  Set CrHNJS=MyFile.CreateTextFile(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js")
					  CrHNJS.write JSCodeStr
				  Set MyFile=nothing
				  ClsJSObj.Close
				  Set ClsJSObj = Nothing
			   else
				  PCssD = JSCodeStr
			   end if
			Else
				PCssD = "生成JS文件时未找到参数"
			End If
	End Function

	Public Function PCssE(EName,CreateFileTF)
		Dim ClsJSObj,ClsJSFileObj,ClsFileSql,ClsNewsObj,TempEName,JSCodeStr,i,MyFile,CrHNJS,OpenMode
		Dim NewsLinkStr
		Set ClsJSObj = Conn.Execute("Select * From FS_FreeJS where EName='"&EName&"'")
			If Not ClsJSObj.eof then
			  JSCodeStr = "document.write('<table class="""&ClsJSObj("BackCSS")&""" width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr>"
			  Set ClsJSFileObj=server.createobject(G_FS_RS)
			  ClsFileSql="select * from FS_FreeJsFile where JSName='"&EName&"' and DelFlag=0 order by ToJsTime desc"
			  ClsJSFileObj.open ClsFileSql,Conn,1,1
			  If ClsJSFileObj.eof then 
				JSCodeStr = JSCodeStr & "<td>此JS内暂无新闻</td></tr>"
			  end if
			  If ClsJSObj("OpenMode")=1 then
				  OpenMode = "target=_blank"
			  Else
				  OpenMode = "target=_self"
			  End If
				ListSpaceStr = ""
				for Temp_i = 1 to Cint(ClsJSObj("RowSpace")\2)
					ListSpaceStr = ListSpaceStr & "&nbsp;"
				next 
			  for i=1 to ClsJSObj("NewsNum")
				  If ClsJSFileObj.eof then Exit For
				  Set ClsNewsObj = Conn.Execute("Select * From FS_News where FileName='"&ClsJSFileObj("FileName")&"'")
				  NewsLinkStr = GetOneNewsLinkURL(ClsNewsObj("NewsID"))
				  JSCodeStr = JSCodeStr & "<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td rowspan=""2"">"&ListSpaceStr&"</td><td valign=""top"" align=""center""><a href="&NewsLinkStr&" "&OpenMode&"><img src="& AvailableDoMain & ClsJSFileObj("PicPath") &" height="&ClsJSObj("PicHeight")&" width="&ClsJSObj("PicWidth")&" border=""0""></a></td><td rowspan=""2"">"&ListSpaceStr&"</td></tr>"
				  If ClsJSObj("MoreContent")=1 then
					  JSCodeStr = JSCodeStr & "<tr><td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......<br><div align=""right""><a class="""&ClsJSObj("LinkCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&ClsJSObj("LinkWord")&"</a></div></td></tr></table></td>"
				  Else
					  JSCodeStr = JSCodeStr & "<tr><td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......</td></tr></table></td>"
				  End If
				  ClsNewsObj.Close
				  Set ClsNewsObj = Nothing
				  ClsJSFileObj.MoveNext
				  if i mod Cint(ClsJSObj("RowNum")) = 0 or ClsJSFileObj.eof then
					  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))&""" height="""&ClsJSObj("RowSpace")&""" background="""&ClsJSObj("RowBettween")&"""></td></tr><tr>"
				  end if
			  next 
			  ClsJSFileObj.Close 
			  Set ClsJSFileObj = Nothing 
			  JSCodeStr = JSCodeStr & "</tr></table>');"
			  JSCodeStr = Replace(JSCodeStr,"<tr></tr>","")
			  JSCodeStr = Replace(Replace(JSCodeStr,Chr(13),""),Chr(10),"")
			  if CreateFileTF = True then
				  Set MyFile=Server.CreateObject(G_FS_FSO)
				  If MyFile.FileExists(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js") then
					 MyFile.DeleteFile(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js")
				  End if
				  Set CrHNJS=MyFile.CreateTextFile(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js")
					  CrHNJS.write JSCodeStr
				  Set MyFile=nothing
				  ClsJSObj.Close
				  Set ClsJSObj = Nothing
			   else
				  PCssE = JSCodeStr
			   end if
			Else
				PCssE = "生成JS文件时未找到参数"
			End If
	End Function

	Public Function PCssF(EName,CreateFileTF)
		Dim ClsJSObj,ClsJSFileObj,ClsFileSql,ClsNewsObj,TempEName,JSCodeStr,i,MyFile,CrHNJS,OpenMode
		Dim NewsLinkStr
		Set ClsJSObj = Conn.Execute("Select * From FS_FreeJS where EName='"&EName&"'")
			If Not ClsJSObj.eof then
			  JSCodeStr = "document.write('<table class="""&ClsJSObj("BackCSS")&""" width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr>"
			  Set ClsJSFileObj=server.createobject(G_FS_RS)
			  ClsFileSql="select * from FS_FreeJsFile where JSName='"&EName&"' and DelFlag=0 order by ToJsTime desc"
			  ClsJSFileObj.open ClsFileSql,Conn,1,1
			  If ClsJSFileObj.eof then 
				JSCodeStr = JSCodeStr & "<td>此JS内暂无新闻</td></tr>"
			  end if
			  If ClsJSObj("OpenMode")=1 then
				  OpenMode = "target=_blank"
			  Else
				  OpenMode = "target=_self"
			  End If
				ListSpaceStr = ""
				for Temp_i = 1 to Cint(ClsJSObj("RowSpace"))
					ListSpaceStr = ListSpaceStr & "&nbsp;"
				next 
			  for i=1 to ClsJSObj("NewsNum")
				  If ClsJSFileObj.eof then Exit For
				  Set ClsNewsObj = Conn.Execute("Select * From FS_News where FileName='"&ClsJSFileObj("FileName")&"'")
				  NewsLinkStr = GetOneNewsLinkURL(ClsNewsObj("NewsID"))
				  JSCodeStr = JSCodeStr & "<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td valign=""top"" align=""center""><a href="&NewsLinkStr&" "&OpenMode&"><img src="& AvailableDoMain & ClsJSFileObj("PicPath") &" height="&ClsJSObj("PicHeight")&" width="&ClsJSObj("PicWidth")&" border=""0""></a></td>"
				  If ClsJSObj("MoreContent")=1 then
					  JSCodeStr = JSCodeStr & "<td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......<br><div align=""right""><a class="""&ClsJSObj("LinkCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&ClsJSObj("LinkWord")&"</a></div></td><td>"&ListSpaceStr&"</td></tr></table></td>"
				  Else
					  JSCodeStr = JSCodeStr & "<td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......</td><td>"&ListSpaceStr&"</td></tr></table></td>"
				  End If
				  ClsNewsObj.Close
				  Set ClsNewsObj = Nothing
				  ClsJSFileObj.MoveNext
				  if i mod Cint(ClsJSObj("RowNum")) = 0 or ClsJSFileObj.eof then
					  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))&""" height="""&ClsJSObj("RowSpace")&""" background="""&ClsJSObj("RowBettween")&"""></td></tr>"
				  end if
			  next 
			  ClsJSFileObj.Close 
			  Set ClsJSFileObj = Nothing 
			  JSCodeStr = JSCodeStr & "</table>');"
			  JSCodeStr = Replace(Replace(JSCodeStr,Chr(13),""),Chr(10),"")
			  if CreateFileTF = True then
				  Set MyFile=Server.CreateObject(G_FS_FSO)
				  If MyFile.FileExists(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js") then
					 MyFile.DeleteFile(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js")
				  End if
				  Set CrHNJS=MyFile.CreateTextFile(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js")
					  CrHNJS.write JSCodeStr
				  Set MyFile=nothing
				  ClsJSObj.Close
				  Set ClsJSObj = Nothing
			   else
				  PCssF = JSCodeStr
			   end if
			Else
				PCssF = "生成JS文件时未找到参数"
			End If
	End Function

	Public Function PCssG(EName,CreateFileTF)
		Dim ClsJSObj,ClsJSFileObj,ClsFileSql,ClsNewsObj,TempEName,JSCodeStr,i,MyFile,CrHNJS,OpenMode
		Dim NewsLinkStr
		Set ClsJSObj = Conn.Execute("Select * From FS_FreeJS where EName='"&EName&"'")
			If Not ClsJSObj.eof then
			  JSCodeStr = "document.write('<table class="""&ClsJSObj("BackCSS")&""" width=""100%"" border=""0"" cellpadding=""0"" cellspacing="""&ClsJSObj("RowSpace")&"""><tr>"
			  Set ClsJSFileObj=server.createobject(G_FS_RS)
			  ClsFileSql="select * from FS_FreeJsFile where JSName='"&EName&"' and DelFlag=0 order by ToJsTime desc"
			  ClsJSFileObj.open ClsFileSql,Conn,1,1
			  If ClsJSFileObj.eof then 
				JSCodeStr = JSCodeStr & "<td>此JS内暂无新闻</td></tr>"
			  end if
			  If ClsJSObj("OpenMode")=1 then
				  OpenMode = "target=_blank"
			  Else
				  OpenMode = "target=_self"
			  End If
			  JSCodeStr = JSCodeStr & "<td valign=""top"" align=""center""><img src="& AvailableDoMain & ClsJSObj("PicPath") &" height="&ClsJSObj("PicHeight")&" width="&ClsJSObj("PicWidth")&" border=""0""></td><td><table width=""100%"" border=""0"" cellpadding=""0"" cellspacing="""&ClsJSObj("RowSpace")&"""><tr>"
			  for i=1 to ClsJSObj("NewsNum")
				  If ClsJSFileObj.eof then Exit For
				  Set ClsNewsObj = Conn.Execute("Select * From FS_News where FileName='"&ClsJSFileObj("FileName")&"'")
				  NewsLinkStr = GetOneNewsLinkURL(ClsNewsObj("NewsID"))
				  If ClsJSObj("ShowTimeTF")="1" then
					  JSCodeStr = JSCodeStr &"<td>"&ClsJSObj("NaviPic")&"<a class="""&ClsJSObj("TitleCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&GotTopic(ClsNewsObj("Title"),ClsJSObj("NewsTitleNum"))&"</a></td><td><Span class="""&ClsJSObj("DateCSS")&""">"&DateFormat(ClsNewsObj("AddDate"),""&ClsJSObj("DateType")&"")&"</Span></td>"
				  Else
					  JSCodeStr = JSCodeStr &"<td>"&ClsJSObj("NaviPic")&"<a class="""&ClsJSObj("TitleCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&GotTopic(ClsNewsObj("Title"),ClsJSObj("NewsTitleNum"))&"</a></td>"
				  End If
				  ClsNewsObj.Close
				  Set ClsNewsObj = Nothing
				  ClsJSFileObj.MoveNext
				  if i mod Cint(ClsJSObj("RowNum")) = 0 and not ClsJSFileObj.eof then
					If ClsJSObj("ShowTimeTF")="1" then
					  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))*2&""" height="""&ClsJSObj("RowSpace")&""" background="""&ClsJSObj("RowBettween")&"""></td></tr><tr>"
					Else
					  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))&""" height="""&ClsJSObj("RowSpace")&""" background="""&ClsJSObj("RowBettween")&"""></td></tr><tr>"
					End If
				  end if
				  if ClsJSFileObj.eof then
					If ClsJSObj("ShowTimeTF")="1" then
					  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))*2&""" height="""&ClsJSObj("RowSpace")&""" background="""&ClsJSObj("RowBettween")&"""></td></tr></table></td></tr>"
					Else
					  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))&""" height="""&ClsJSObj("RowSpace")&""" background="""&ClsJSObj("RowBettween")&"""></td></tr></table></td></tr>"
					End If
				  end if
			  next 
			  ClsJSFileObj.Close 
			  Set ClsJSFileObj = Nothing 
			  JSCodeStr = JSCodeStr & "</table>');"
			  JSCodeStr = Replace(Replace(JSCodeStr,Chr(13),""),Chr(10),"")
			  if CreateFileTF = True then
				  Set MyFile=Server.CreateObject(G_FS_FSO)
				  If MyFile.FileExists(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js") then
					 MyFile.DeleteFile(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js")
				  End if
				  Set CrHNJS=MyFile.CreateTextFile(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js")
					  CrHNJS.write JSCodeStr
				  Set MyFile=nothing
				  ClsJSObj.Close
				  Set ClsJSObj = Nothing
			   else
				  PCssG = JSCodeStr
			   end if
			Else
				PCssG = "生成JS文件时未找到参数"
			End If
	End Function

	Public Function PCssH(EName,CreateFileTF)
		Dim ClsJSObj,ClsJSFileObj,ClsFileSql,ClsNewsObj,TempEName,JSCodeStr,i,MyFile,CrHNJS,OpenMode
		Dim NewsLinkStr
		Set ClsJSObj = Conn.Execute("Select * From FS_FreeJS where EName='"&EName&"'")
			If Not ClsJSObj.eof then
			  JSCodeStr = "document.write('<table class="""&ClsJSObj("BackCSS")&""" width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr>"
			  Set ClsJSFileObj=server.createobject(G_FS_RS)
			  ClsFileSql="select * from FS_FreeJsFile where JSName='"&EName&"' and DelFlag=0 order by ToJsTime desc"
			  ClsJSFileObj.open ClsFileSql,Conn,1,1
			  If ClsJSFileObj.eof then 
				JSCodeStr = JSCodeStr & "<td>此JS内暂无新闻</td>"
			  end if
			  If ClsJSObj("OpenMode")=1 then
				  OpenMode = "target=_blank"
			  Else
				  OpenMode = "target=_self"
			  End If
				ListSpaceStr = ""
				for Temp_i = 1 to Cint(ClsJSObj("RowSpace"))
					ListSpaceStr = ListSpaceStr & "&nbsp;"
				next 
			  for i=1 to ClsJSObj("NewsNum")
				  If ClsJSFileObj.eof then Exit For
				  Set ClsNewsObj = Conn.Execute("Select * From FS_News where FileName='"&ClsJSFileObj("FileName")&"'")
				  NewsLinkStr = GetOneNewsLinkURL(ClsNewsObj("NewsID"))
				  JSCodeStr = JSCodeStr & "<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td valign=""top"" align=""left""><a href="&NewsLinkStr&" "&OpenMode&"><img src="& AvailableDoMain & ClsJSFileObj("PicPath") &" height="&ClsJSObj("PicHeight")&" width="&ClsJSObj("PicWidth")&" border=""0""></a></td>"
				  If ClsJSObj("ShowTimeTF")="1" then
					  JSCodeStr = JSCodeStr & "<td><div align=""left"">"&ClsJSObj("NaviPic")&"<a class="""&ClsJSObj("TitleCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&GotTopic(ClsNewsObj("Title"),ClsJSObj("NewsTitleNum"))&"</a>&nbsp;<Span class="""&ClsJSObj("DateCSS")&""">"&DateFormat(ClsNewsObj("AddDate"),""&ClsJSObj("DateType")&"")&"</Span></div></td><td rowspan=""2"">"&ListSpaceStr&"</td></tr>"
				  Else
					  JSCodeStr = JSCodeStr & "<td><div align=""left"">"&ClsJSObj("NaviPic")&"<a class="""&ClsJSObj("TitleCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&GotTopic(ClsNewsObj("Title"),ClsJSObj("NewsTitleNum"))&"</a></div></td><td rowspan=""2"">"&ListSpaceStr&"</td></tr>"
				  End If
				  If ClsJSObj("MoreContent")=1 then
					  JSCodeStr = JSCodeStr & "<tr><td colspan=""2""><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......<br><div align=""right""><a class="""&ClsJSObj("LinkCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&ClsJSObj("LinkWord")&"</a></div></td></tr></table></td>"
				  Else
					  JSCodeStr = JSCodeStr & "<tr><td colspan=""2""><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......</td></tr></table></td>"
				  End If
				  ClsNewsObj.Close
				  Set ClsNewsObj = Nothing
				  ClsJSFileObj.MoveNext
				  if i mod Cint(ClsJSObj("RowNum")) = 0 or ClsJSFileObj.eof then
					  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))&""" height="""&ClsJSObj("RowSpace")&""" background="""&ClsJSObj("RowBettween")&"""></td></tr><tr>"
				  end if
			  next 
			  ClsJSFileObj.Close 
			  Set ClsJSFileObj = Nothing 
			  JSCodeStr = JSCodeStr & "</tr></table>');"
			  JSCodeStr = Replace(JSCodeStr,"<tr></tr>","")
			  JSCodeStr = Replace(Replace(JSCodeStr,Chr(13),""),Chr(10),"")
			  if CreateFileTF = True then
				  Set MyFile=Server.CreateObject(G_FS_FSO)
				  If MyFile.FileExists(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js") then
					 MyFile.DeleteFile(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js")
				  End if
				  Set CrHNJS=MyFile.CreateTextFile(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js")
					  CrHNJS.write JSCodeStr
				  Set MyFile=nothing
				  ClsJSObj.Close
				  Set ClsJSObj = Nothing
			   else
				  PCssH = JSCodeStr
			   end if
			Else
				PCssH = "生成JS文件时未找到参数"
			End If
	End Function

	Public Function PCssI(EName,CreateFileTF)
		Dim ClsJSObj,ClsJSFileObj,ClsFileSql,ClsNewsObj,TempEName,JSCodeStr,i,MyFile,CrHNJS,OpenMode
		Dim NewsLinkStr
		Set ClsJSObj = Conn.Execute("Select * From FS_FreeJS where EName='"&EName&"'")
			If Not ClsJSObj.eof then
			  JSCodeStr = "document.write('<table class="""&ClsJSObj("BackCSS")&""" width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr>"
			  Set ClsJSFileObj=server.createobject(G_FS_RS)
			  ClsFileSql="select * from FS_FreeJsFile where JSName='"&EName&"' and DelFlag=0 order by ToJsTime desc"
			  ClsJSFileObj.open ClsFileSql,Conn,1,1
			  If ClsJSFileObj.eof then 
				JSCodeStr = JSCodeStr & "<td>此JS内暂无新闻</td></tr>"
			  end if
			  If ClsJSObj("OpenMode")=1 then
				  OpenMode = "target=_blank"
			  Else
				  OpenMode = "target=_self"
			  End If
				ListSpaceStr = ""
				for Temp_i = 1 to Cint(ClsJSObj("RowSpace")\2)
					ListSpaceStr = ListSpaceStr & "&nbsp;"
				next 
			  for i=1 to ClsJSObj("NewsNum")
				  If ClsJSFileObj.eof then Exit For
				  Set ClsNewsObj = Conn.Execute("Select * From FS_News where FileName='"&ClsJSFileObj("FileName")&"'")
				  NewsLinkStr = GetOneNewsLinkURL(ClsNewsObj("NewsID"))
				  If ClsJSObj("ShowTimeTF")="1" then
					  JSCodeStr = JSCodeStr & "<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td rowspan=""2"">"&ListSpaceStr&"</td><td colspan=""3""><div align=""center"">"&ClsJSObj("NaviPic")&"<a class="""&ClsJSObj("TitleCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&GotTopic(ClsNewsObj("Title"),ClsJSObj("NewsTitleNum"))&"</a>&nbsp;<Span class="""&ClsJSObj("DateCSS")&""">"&DateFormat(ClsNewsObj("AddDate"),""&ClsJSObj("DateType")&"")&"</Span></div></td><td rowspan=""2"">"&ListSpaceStr&"</td></tr>"
				  Else
					  JSCodeStr = JSCodeStr & "<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td rowspan=""2"">"&ListSpaceStr&"</td><td colspan=""3""><div align=""center"">"&ClsJSObj("NaviPic")&"<a class="""&ClsJSObj("TitleCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&GotTopic(ClsNewsObj("Title"),ClsJSObj("NewsTitleNum"))&"</a></div></td><td rowspan=""2"">"&ListSpaceStr&"</td></tr>"
				  End If
				  JSCodeStr = JSCodeStr & "<tr><td valign=""top""><a href="&NewsLinkStr&" "&OpenMode&"><img src="& AvailableDoMain & ClsJSFileObj("PicPath") &" height="&ClsJSObj("PicHeight")&" width="&ClsJSObj("PicWidth")&" border=""0""></a></td><td>&nbsp;</td>"			  
				  If ClsJSObj("MoreContent")=1 then
					  JSCodeStr = JSCodeStr & "<td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......<br><div align=""right""><a class="""&ClsJSObj("LinkCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&ClsJSObj("LinkWord")&"</a></div></td></tr></table></td>"
				  Else
					  JSCodeStr = JSCodeStr & "<td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......</td></tr></table></td>"
				  End If
				  ClsNewsObj.Close
				  Set ClsNewsObj = Nothing
				  ClsJSFileObj.MoveNext
				  if i mod Cint(ClsJSObj("RowNum")) = 0 or ClsJSFileObj.eof then
					  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))&""" height="""&ClsJSObj("RowSpace")&""" background="""&ClsJSObj("RowBettween")&"""></td></tr><tr>"
				  end if
			  next 
			  ClsJSFileObj.Close 
			  Set ClsJSFileObj = Nothing 
			  JSCodeStr = JSCodeStr & "</tr></table>');"
			  JSCodeStr = Replace(JSCodeStr,"<tr></tr>","")
			  JSCodeStr = Replace(Replace(JSCodeStr,Chr(13),""),Chr(10),"")
			  if CreateFileTF = True then
				  Set MyFile=Server.CreateObject(G_FS_FSO)
				  If MyFile.FileExists(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js") then
					 MyFile.DeleteFile(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js")
				  End if
				  Set CrHNJS=MyFile.CreateTextFile(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js")
					  CrHNJS.write JSCodeStr
				  Set MyFile=nothing
				  ClsJSObj.Close
				  Set ClsJSObj = Nothing
			   else
				  PCssI = JSCodeStr
			   end if
			Else
				PCssI = "生成JS文件时未找到参数"
			End If
	End Function

	Public Function PCssJ(EName,CreateFileTF)
		Dim ClsJSObj,ClsJSFileObj,ClsFileSql,ClsNewsObj,TempEName,JSCodeStr,i,MyFile,CrHNJS,OpenMode
		Dim NewsLinkStr
		Set ClsJSObj = Conn.Execute("Select * From FS_FreeJS where EName='"&EName&"'")
			If Not ClsJSObj.eof then
			  JSCodeStr = "document.write('<table class="""&ClsJSObj("BackCSS")&""" width=100% border=0 cellpadding=0 cellspacing="""&ClsJSObj("RowSpace")&"""><tr>"
			  Set ClsJSFileObj=server.createobject(G_FS_RS)
			  ClsFileSql="select * from FS_FreeJsFile where JSName='"&EName&"' and DelFlag=0 order by ToJsTime desc"
			  ClsJSFileObj.open ClsFileSql,Conn,1,1
			  If ClsJSFileObj.eof then 
				JSCodeStr = JSCodeStr & "<td>此JS内暂无新闻</td>"
			  end if
			  If ClsJSObj("OpenMode")=1 then
				  OpenMode = "target=_blank"
			  Else
				  OpenMode = "target=_self"
			  End If
			  for i=1 to ClsJSObj("NewsNum")
				  If ClsJSFileObj.eof then Exit For
				  Set ClsNewsObj = Conn.Execute("Select * From FS_News where FileName='"&ClsJSFileObj("FileName")&"'")
				  NewsLinkStr = GetOneNewsLinkURL(ClsNewsObj("NewsID"))
				  If ClsJSObj("ShowTimeTF")="1" then
					  JSCodeStr = JSCodeStr &"<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table background="""& AvailableDoMain & ClsJSFileObj("PicPath")&""" width=""100%"" border=""0"" cellpadding=""0"" cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td>"& ClsJSObj("NaviPic") &"<a class="""&ClsJSObj("TitleCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&GotTopic(ClsNewsObj("Title"),ClsJSObj("NewsTitleNum"))&"</a>&nbsp;<Span class="""&ClsJSObj("DateCSS")&""">"&DateFormat(ClsNewsObj("AddDate"),""&ClsJSObj("DateType")&"")&"</Span></td></tr>"
				  Else
					  JSCodeStr = JSCodeStr &"<td width="&Cint(100/Cint(ClsJSObj("RowNum")))&"% valign=""top""><table background="""& AvailableDoMain & ClsJSFileObj("PicPath")&""" width=""100%"" border=""0"" cellpadding=""0"" cellspacing="""&ClsJSObj("RowSpace")&"""><tr><td>"& ClsJSObj("NaviPic") &"<a class="""&ClsJSObj("TitleCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&GotTopic(ClsNewsObj("Title"),ClsJSObj("NewsTitleNum"))&"</a></td></tr>"
				  End If
				  If ClsJSObj("MoreContent")=1 then
					  JSCodeStr = JSCodeStr & "<tr><td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......<br><div align=""right""><a class="""&ClsJSObj("LinkCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&ClsJSObj("LinkWord")&"</a></div></td></tr></table></td>"
				  Else
					  JSCodeStr = JSCodeStr & "<tr><td><Span class="""&ClsJSObj("ContentCSS")&""">"&TitleSpaceStr&GotTopic(Replace(Replace(Replace(LoseHtml(ClsNewsObj("Content")),chr(13) & chr(10),""),"[Page]",""),"&nbsp;",""),ClsJSObj("ContentNum"))&"</Span>......</td></tr></table></td>"
				  End If
				  ClsNewsObj.Close
				  Set ClsNewsObj = Nothing
				  ClsJSFileObj.MoveNext
				  if i mod Cint(ClsJSObj("RowNum")) = 0 or ClsJSFileObj.eof then
					  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))&""" height="""&ClsJSObj("RowSpace")&""" background="""&ClsJSObj("RowBettween")&"""></td></tr><tr>"
				  end if
			  next 
			  ClsJSFileObj.Close 
			  Set ClsJSFileObj = Nothing 
			  JSCodeStr = JSCodeStr & "</tr></table>');"
			  JSCodeStr = Replace(JSCodeStr,"<tr></tr>","")
			  JSCodeStr = Replace(Replace(JSCodeStr,Chr(13),""),Chr(10),"")
			  if CreateFileTF = True then
				  Set MyFile=Server.CreateObject(G_FS_FSO)
				  If MyFile.FileExists(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js") then
					 MyFile.DeleteFile(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js")
				  End if
				  Set CrHNJS=MyFile.CreateTextFile(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js")
					  CrHNJS.write JSCodeStr
				  Set MyFile=nothing
				  ClsJSObj.Close
				  Set ClsJSObj = Nothing
			   else
				  PCssJ = JSCodeStr
			   end if
			Else
				PCssJ = "生成JS文件时未找到参数"
			End If
	End Function

	Public Function PCssK(EName,CreateFileTF)
		Dim ClsJSObj,ClsJSFileObj,ClsFileSql,ClsNewsObj,TempEName,JSCodeStr,i,MyFile,CrHNJS,OpenMode
		Dim NewsLinkStr
		Set ClsJSObj = Conn.Execute("Select * From FS_FreeJS where EName='"&EName&"'")
			If Not ClsJSObj.eof then
			  JSCodeStr = "document.write('<table background="""& AvailableDoMain & ClsJSObj("PicPath")&""" width=""100%"" border=""0"" cellpadding=""0"" cellspacing="""&ClsJSObj("RowSpace")&"""><tr>"
			  Set ClsJSFileObj=server.createobject(G_FS_RS)
			  ClsFileSql="select * from FS_FreeJsFile where JSName='"&EName&"' and DelFlag=0 order by ToJsTime desc"
			  ClsJSFileObj.open ClsFileSql,Conn,1,1
			  If ClsJSFileObj.eof then
				JSCodeStr = JSCodeStr & "<td>此JS内暂无新闻</td>"
			  End If
			  If ClsJSObj("OpenMode")=1 then
				  OpenMode = "target=_blank"
			  Else
				  OpenMode = "target=_self"
			  End If
			  for i=1 to ClsJSObj("NewsNum")
				  If ClsJSFileObj.eof then Exit For
				  Set ClsNewsObj = Conn.Execute("Select * From FS_News where FileName='"&ClsJSFileObj("FileName")&"'")
				  NewsLinkStr = GetOneNewsLinkURL(ClsNewsObj("NewsID"))
				  If ClsJSObj("ShowTimeTF")="1" then
					  JSCodeStr = JSCodeStr &"<td valign=middle>"&ClsJSObj("NaviPic")&"<a class="""&ClsJSObj("TitleCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&GotTopic(ClsNewsObj("Title"),ClsJSObj("NewsTitleNum"))&"</a></td><td><Span class="""&ClsJSObj("DateCSS")&""">"&DateFormat(ClsNewsObj("AddDate"),""&ClsJSObj("DateType")&"")&"</Span></td>"
				  Else
					  JSCodeStr = JSCodeStr &"<td valign=middle>"&ClsJSObj("NaviPic")&"<a class="""&ClsJSObj("TitleCSS")&""" href="&NewsLinkStr&" "&OpenMode&">"&GotTopic(ClsNewsObj("Title"),ClsJSObj("NewsTitleNum"))&"</a></td>"
				  End If
				  ClsNewsObj.Close
				  Set ClsNewsObj = Nothing
				  ClsJSFileObj.MoveNext
				  if i mod Cint(ClsJSObj("RowNum")) = 0 or ClsJSFileObj.eof then
					if ClsJSObj("ShowTimeTF")=1 then
						  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))*2&""" height="""&ClsJSObj("RowSpace")&""" background="""&ClsJSObj("RowBettween")&"""></td></tr><tr>"
					else
						  JSCodeStr = JSCodeStr &"</tr><tr><td colspan="""&Cint(ClsJSObj("RowNum"))&""" height="""&ClsJSObj("RowSpace")&""" background="""&ClsJSObj("RowBettween")&"""></td></tr><tr>"
					end if
				  end if
			  next 
			  ClsJSFileObj.Close 
			  Set ClsJSFileObj = Nothing 
			  JSCodeStr = JSCodeStr & "</tr></table>');"
			  JSCodeStr = Replace(JSCodeStr,"<tr></tr>","")
			  JSCodeStr = Replace(Replace(JSCodeStr,Chr(13),""),Chr(10),"")
			  if CreateFileTF = True then
				  Set MyFile=Server.CreateObject(G_FS_FSO)
				  If MyFile.FileExists(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js") then
					 MyFile.DeleteFile(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js")
				  End if
				  Set CrHNJS=MyFile.CreateTextFile(Server.MapPath(TempSysRootDir&"\JS\FreeJs")&"\"& EName &".js")
					  CrHNJS.write JSCodeStr
				  Set MyFile=nothing
				  ClsJSObj.Close
				  Set ClsJSObj = Nothing
			   else
				  PCssK = JSCodeStr
			   end if
			Else
				PCssK = "生成JS文件时未找到参数"
			End If
	End Function
End Class
%>