<% Option Explicit %>
<%
Response.Buffer = true
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache"
%>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="inc/Function.asp" -->
<!--#include file="inc/Config.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
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
Dim DBC,Conn,CollectConn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
DBC.ConnStr = CollectDBConnectionStr
Set CollectConn = DBC.OpenConnection()
Set DBC = Nothing

%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P080106") then Call ReturnError1()
Dim NewsListPagesNumber
'////////////////////////////////////////////////////////////////////////////////////
'其他地方的变量请不要设置，如果设置了，有可能引起不能够正常采集。
Server.ScriptTimeOut = 10000    '设置脚本超时 
NewsListPagesNumber = 300         '采集新闻列表分页最大数量
'////////////////////////////////////////////////////////////////////////////////////
Dim AvailableDoMain '站点域名--配置信息 
Dim DummyDir '虚拟目录--配置信息 
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
	If SysRootDir <> "" then
		DummyDir = "/" & SysRootDir
	Else
		DummyDir = ""
	End If
End Sub

Dim SiteID,ErrorInfoStr,Action
Action = Request("Action")
SiteID = Request("SiteID")
ErrorInfoStr = ""

Dim SysClassID,SaveIMGPath
Dim ListHeadSetting,ListFootSetting
Dim LinkHeadSetting,LinkFootSetting
Dim PagebodyHeadSetting,PagebodyFootSetting
Dim PageTitleHeadSetting,PageTitleFootSetting
Dim OtherPageFootSetting,OtherPageHeadSetting
Dim OtherNewsPageHeadSetting,OtherNewsPageFootSetting
Dim AuthorHeadSetting,AuthorFootSetting
Dim SourceHeadSetting,SourceFootSetting
Dim AddDateHeadSetting,AddDateFootSetting
Dim IndexRule,StartPageNum,EndPageNum,HandPageContent,OtherType
Dim IsStyle,IsDiv,IsA,IsClass,IsFont,IsSpan,IsObjectTF,IsIFrame,IsScript
Dim HandSetAuthor,HandSetSource,HandSetAddDate
Dim TextTF,SaveRemotePic,IsReverse
Dim SysTemplet '新闻模板
Dim ObjURL
Dim ReturnValue
Dim CollectStartLocation
Dim CollectEndFlag
CollectEndFlag = False
Dim CollectObjURL 
Dim CollectedPageURL
CollectedPageURL = Request("CollectedPageURL")
Dim SiteName
Dim CollectingSiteID
Dim CollectSiteIndex
Dim AllNewsNumber,CollectOKNumber
AllNewsNumber = Request("AllNewsNumber")
if AllNewsNumber = "" then
	AllNewsNumber = 0
else
	AllNewsNumber = CLng(AllNewsNumber)
end if
CollectOKNumber = Request("CollectOKNumber")
if CollectOKNumber = "" then
	CollectOKNumber = 0
else
	CollectOKNumber = CLng(CollectOKNumber)
end if
CollectSiteIndex = Request("CollectSiteIndex")
if CollectSiteIndex = "" then
	CollectSiteIndex = 0
else
	CollectSiteIndex = CInt(CollectSiteIndex)
end if
Dim CollectPageNumber
CollectPageNumber = Request("CollectPageNumber")
if CollectPageNumber = "" then
	CollectPageNumber = 0
else
	CollectPageNumber = CInt(CollectPageNumber)
end if
CollectStartLocation = Request("CollectStartLocation")
if CollectStartLocation = "" then CollectStartLocation = 0
Dim Num
Num = Request("Num")
If Num = "allNews" Or Num="" Then 
	Num = 0
Else
	Num = CInt(Num)
End If
Dim CollectType
CollectType = Request("CollectType")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>正在采集</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<body topmargin="2" leftmargin="2" oncontextmenu="//return false;">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="35" id="StopCollect" align="center" alt="停止采集" onClick="location.href='Site.asp';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">取消</td>
		  <td width=2 class="Gray">|</td>
          <td width="35" id="SaveCollect" align="center" alt="保存采集进度并返回" onClick="location.href='Site.asp';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td>
		  <td width=2 class="Gray">|</td>
		  <td width="35" align="center" alt="后退" onClick="history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
          <td>&nbsp;</td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="20"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
		<%If CollectType="ResumeCollect" then%>
			<td width="50%;" align="right"><font color="#FF0000" id="CollectEndArea">正在续采</font></td>
		<%else%>
			<td width="50%;" align="right"><font color="#FF0000" id="CollectEndArea">正在采集</font></td>
		<%End if%>
			<td width="50%;">&nbsp;<font color="#FF0000" id="ShowInfoArea" size="+1">&nbsp;</font></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td valign="middle">
<%
if Action = "Submit" then
	if SiteID <> "" then
		GetCollectPara
		If AllNewsNumber>=Num And Num<>0 Then 
			CollectEndFlag = True
		End If
		if CollectEndFlag then
			if ErrorInfoStr <> "" then
				Response.Write(ErrorInfoStr)
			else
				ReturnValue = "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong>采集结束</strong>： 共读取" & AllNewsNumber & "条新闻，采集成功" & CollectOKNumber & "条新闻。"
				Response.Write(ReturnValue)
				Response.Write("<script language=""JavaScript"">setTimeout('SetCollectEndStr();',100);</script>")
			end if
		elseif CollectType<>"ResumeCollect" Then
			GetNewsPageContent()
			if CollectStartLocation = 0 then
				ReturnValue = "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong><font color=red>采集分页" & CollectPageNumber & "</font></strong>：" & "<a target=""_blank"" href=""" & ObjURL & """>" & ObjURL & "</a><br>" & ReturnValue
			else
				ReturnValue = "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong><font color=red>采集分页" & CollectPageNumber + 1 & "</font></strong>：" & "<a target=""_blank"" href=""" & ObjURL & """>" & ObjURL & "</a><br>" & ReturnValue
			end if
			ReturnValue = "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong><font color=red>采集站点</font></strong>：" & SiteName & "<br>" & ReturnValue
			ReturnValue = "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong><font color=red>采集结果</font></strong>：已经读取" & AllNewsNumber & "条新闻，保存" & CollectOKNumber & "条新闻<br>" & ReturnValue
			Response.Write(ReturnValue & "<meta http-equiv=""refresh"" content=""2;url=Collecting.asp?Action=Submit&CollectPageNumber=" & CollectPageNumber & "&SiteID=" & SiteID & "&CollectStartLocation=" & CollectStartLocation & "&CollectedPageURL=" & CollectedPageURL & "&CollectSiteIndex=" & CollectSiteIndex & "&Num=" & Num & "&AllNewsNumber=" & AllNewsNumber & "&CollectOKNumber=" & CollectOKNumber & """>")
		else
			ResumeGetNewsPageContent()
			if CollectStartLocation = 0 then
				ReturnValue = "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong><font color=red>采集分页" & CollectPageNumber & "</font></strong>：" & "<a target=""_blank"" href=""" & ObjURL & """>" & ObjURL & "</a><br>" & ReturnValue
			else
				ReturnValue = "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong><font color=red>采集分页" & CollectPageNumber + 1 & "</font></strong>：" & "<a target=""_blank"" href=""" & ObjURL & """>" & ObjURL & "</a><br>" & ReturnValue
			end if
			ReturnValue = "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong><font color=red>采集站点</font></strong>：" & SiteName & "<br>" & ReturnValue
			ReturnValue = "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong><font color=red>采集结果</font></strong>：已经读取" & AllNewsNumber & "条新闻，续采了" & CollectOKNumber & "条新闻<br>" & ReturnValue
			Response.Write(ReturnValue & "<meta http-equiv=""refresh"" content=""2;url=Collecting.asp?Action=Submit&CollectType=ResumeCollect&CollectPageNumber=" & CollectPageNumber & "&SiteID=" & SiteID & "&CollectStartLocation=" & CollectStartLocation & "&CollectedPageURL=" & CollectedPageURL & "&CollectSiteIndex=" & CollectSiteIndex & "&AllNewsNumber=" & AllNewsNumber & "&CollectOKNumber=" & CollectOKNumber & """>")
		end if
	end if
end if
%>
	</td>
  </tr>
</table>
</body>
</html>
<script language="JavaScript">
var ForwardShow=true;
function ShowPromptInfo()
{
	var TempStr=document.all.ShowInfoArea.innerText;
	if (ForwardShow==true)
	{
		if (TempStr.length>4) ForwardShow=false;
		document.all.ShowInfoArea.innerText=TempStr+'.';
	}
	else
	{
		if (TempStr.length==2) ForwardShow=true;
		document.all.ShowInfoArea.innerText=TempStr.substr(0,TempStr.length-1);
	}
}
function SetCollectEndStr()
{
	document.all.CollectEndArea.innerText='采集结束,3秒钟后返回主页面';
	setTimeout("location='Site.asp';",3000);
}
window.setInterval('ShowPromptInfo()',500);</script>
<% if Action = "" then %>
<script language="JavaScript">
setTimeout("location='?SiteID=<% = SiteID %>&CollectType=<%= CollectType %>&Action=Submit&Num=<%= Num %>';",10);
</script>
<% end if %>
<%
Set Conn = Nothing
Set CollectConn = Nothing
Function GetCollectPara()
	Dim RsSiteObj,Sql,SiteIDArray
	if SiteID = "" then
		ErrorInfoStr = "没有采集站点，请重试"
		Exit Function
	end if
	SiteIDArray = Split(SiteID,"***")
	if CollectSiteIndex > UBound(SiteIDArray) then
		CollectEndFlag = True
		Exit Function
	end if
	CollectingSiteID = SiteIDArray(CollectSiteIndex)
	Sql = "Select * from FS_Site where ID=" & CollectingSiteID
	Set RsSiteObj = CollectConn.Execute(Sql)
	if RsSiteObj.Eof then
		Set RsSiteObj = Nothing
		ErrorInfoStr = "没有采集站点，请重试"
		Exit Function
	else
		SiteName = RsSiteObj("SiteName")
		ListHeadSetting = RsSiteObj("ListHeadSetting")
		ListFootSetting = RsSiteObj("ListFootSetting")
		LinkHeadSetting = RsSiteObj("LinkHeadSetting")
		LinkFootSetting = RsSiteObj("LinkFootSetting")
		PagebodyHeadSetting = RsSiteObj("PagebodyHeadSetting")
		PagebodyFootSetting = RsSiteObj("PagebodyFootSetting")
		PageTitleHeadSetting = RsSiteObj("PageTitleHeadSetting")
		PageTitleFootSetting = RsSiteObj("PageTitleFootSetting")
		OtherPageFootSetting = RsSiteObj("OtherPageFootSetting")
		OtherPageHeadSetting = RsSiteObj("OtherPageHeadSetting")
		OtherNewsPageHeadSetting = RsSiteObj("OtherNewsPageHeadSetting")
		OtherNewsPageFootSetting = RsSiteObj("OtherNewsPageFootSetting")
		AuthorHeadSetting = RsSiteObj("AuthorHeadSetting")
		AuthorFootSetting = RsSiteObj("AuthorFootSetting")
		SourceHeadSetting = RsSiteObj("SourceHeadSetting")
		SourceFootSetting = RsSiteObj("SourceFootSetting")
		AddDateHeadSetting = RsSiteObj("AddDateHeadSetting")
		AddDateFootSetting = RsSiteObj("AddDateFootSetting")
		SysClassID = RsSiteObj("SysClass")
		SysTemplet = RsSiteObj("SysTemplet")
		TextTF = RsSiteObj("TextTF")
		SaveRemotePic = RsSiteObj("SaveRemotePic")
		CollectObjURL = RsSiteObj("objURL")
		SaveIMGPath = RsSiteObj("SaveIMGPath")
		Dim TempSaveIMGPath
		TempSaveIMGPath = SaveIMGPath
		SaveIMGPath =SaveIMGPath &"/"& Year(Date) &"-"& Month(Date) &"/"& Day(Date)
		CreateDateDir(Server.mappath(DummyDir & TempSaveIMGPath))
		IsStyle = RsSiteObj("IsStyle")
		IsDiv = RsSiteObj("IsDiv")
		IsA = RsSiteObj("IsA")
		IsClass = RsSiteObj("IsClass")
		IsFont = RsSiteObj("IsFont")
		IsSpan = RsSiteObj("IsSpan")
		IsObjectTF = RsSiteObj("IsObject")
		IsIFrame = RsSiteObj("IsIFrame")
		IsScript = RsSiteObj("IsScript")
		IndexRule = RsSiteObj("IndexRule")
		StartPageNum = RsSiteObj("StartPageNum")
		EndPageNum = RsSiteObj("EndPageNum")
		HandPageContent = RsSiteObj("HandPageContent")
		OtherType = RsSiteObj("OtherType")
		HandSetAuthor = RsSiteObj("HandSetAuthor")
		HandSetSource = RsSiteObj("HandSetSource")
		HandSetAddDate = RsSiteObj("HandSetAddDate")
		ObjURL = GetOtherURL(CollectPageNumber,RsSiteObj)
		IsReverse=RsSiteObj("IsReverse")
		if ObjURL = "" then
			CollectPageNumber = 0
			CollectStartLocation = 0
			CollectedPageURL = ""
			CollectSiteIndex = CollectSiteIndex + 1
			Set RsSiteObj = Nothing
			GetCollectPara
			Exit Function
		else
			if CollectPageNumber > NewsListPagesNumber then
				CollectPageNumber = 0
				CollectStartLocation = 0
				CollectedPageURL = ""
				CollectSiteIndex = CollectSiteIndex + 1
				Set RsSiteObj = Nothing
				GetCollectPara
				Exit Function
			end if
		end if
	end if
	Set RsSiteObj = Nothing
End Function

Function GetOtherURL(PageNum,Obj) '取得其他新闻列表的URL
	Dim OtherObjURL,OtherResponseAllStr,OtherNewsListArray,i
	if PageNum = 0 then
		GetOtherURL = CollectObjURL
		CollectedPageURL = ""
	else
		Select Case OtherType
			Case 0 '不分页
				GetOtherURL = ""
			Case 1 '标记分页
				if IsNull(OtherPageHeadSetting) OR IsNull(OtherPageFootSetting) OR (OtherPageFootSetting = "") OR (OtherPageHeadSetting = "") then
					GetOtherURL = ""
				else
					if PageNum = 1 then
						CollectedPageURL = CollectObjURL
					end if
					OtherResponseAllStr = GetPageContent(FormatUrl(CollectedPageURL,CollectObjURL))
					OtherObjURL = GetOtherContent(OtherResponseAllStr,OtherPageHeadSetting,OtherPageFootSetting)
					if OtherObjURL <> "" then
						OtherObjURL = FormatUrl(OtherObjURL,CollectObjURL)
					else
						OtherObjURL = ""
					end if
					GetOtherURL = OtherObjURL
				end if
			Case 2 '索引分页
				if IsNull(IndexRule) OR (IndexRule = "") OR IsNull(StartPageNum) OR (StartPageNum = "") OR IsNull(EndPageNum) OR (EndPageNum = "") then
					GetOtherURL = ""
				else
					if Not IsNumeric(StartPageNum) OR Not IsNumeric(EndPageNum) then
						GetOtherURL = ""
					else
						if CInt(StartPageNum) < CInt(EndPageNum) Then '按从小到大的页数
							if PageNum >= CInt(EndPageNum) then
								GetOtherURL = ""
							else
								if PageNum = 1 then
									IndexRule = Replace(FormatUrl(IndexRule,CollectObjURL),"^$^",StartPageNum)
								else
									StartPageNum = CInt(StartPageNum) + PageNum - 1
									IndexRule = Replace(FormatUrl(IndexRule,CollectObjURL),"^$^",StartPageNum)
								end if
								GetOtherURL = IndexRule
							end if
						Else  '按从大到小的页数，从而实现倒序采集，比如从10到1
							if PageNum >= CInt(StartPageNum) then
								GetOtherURL = ""
							else
								if PageNum = 1 then
									IndexRule = Replace(FormatUrl(IndexRule,CollectObjURL),"^$^",StartPageNum)
								else
									EndPageNum = CInt(StartPageNum) - PageNum + 1
									IndexRule = Replace(FormatUrl(IndexRule,CollectObjURL),"^$^",EndPageNum)
								end if
								GetOtherURL = IndexRule
							end if
						end if
					end if
				end if
			Case 3 '手工分页
				if IsNull(HandPageContent) OR (HandPageContent = "") then
					GetOtherURL = ""
				ElseIf InStr(HandPageContent,Chr(10))=0 And PageNum<2 Then
					GetOtherURL = HandPageContent
				Else
					HandPageContent = Split(HandPageContent,Chr(10))
					if PageNum > UBound(HandPageContent) then
						GetOtherURL = ""
					else
						if HandPageContent(PageNum - 1) <> "" then
							GetOtherURL = HandPageContent(PageNum - 1)
						else
							GetOtherURL = ""
						end if
					end if
				end if
			Case Else
				GetOtherURL = ""
		End Select
	end if
End Function

Function GetNewsPageContent()
	Dim NewsPageStr,TitleStr,ContentStr,AuthorStr,SourceStr,AddDate,i
	Dim ResponseAllStr,NewsListStr,NewsLinkStr,RsCheckNewsObj
	Dim NewsListStrArray,TempArray
	ResponseAllStr = GetPageContent(FormatUrl(ObjURL,CollectObjURL))	
	if ResponseAllStr = False then
		CollectPageNumber = CollectPageNumber + 1
		ReturnValue = ReturnValue & "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong>错误</strong>:读取新闻列表页面失败<br>"
		Exit Function
	end if

	Dim BLinkHeadSetting,BLinkFootSetting
	BLinkHeadSetting = False
	BLinkFootSetting = False
	If Instr(LinkHeadSetting,"[变量]")<=0 Then
		BLinkHeadSetting = True
	ElseIf Instr(LinkFootSetting,"[变量]")<=0 Then
		BLinkFootSetting = True
	End If
	If InStr(ResponseAllStr,ListHeadSetting)>0 And InStr(ResponseAllStr,ListFootSetting) Then
		NewsListStr = GetOtherContent(ResponseAllStr,ListHeadSetting,ListFootSetting)
	Else 
		NewsListStr = ResponseAllStr
	End If

	If BLinkHeadSetting Then
		NewsListStr = Mid(NewsListStr,Instr(NewsListStr,LinkHeadSetting)+len(LinkHeadSetting))
		NewsListStrArray = Split(NewsListStr,LinkHeadSetting)
	elseif BLinkFootSetting Then 
		NewsListStr = Left(NewsListStr,InstrRev(NewsListStr,LinkFootSetting))
		NewsListStrArray = Split(NewsListStr,LinkFootSetting)
	End If

	'倒序采集
	If IsReverse="1" then 
		Dim TempArr,j
		TempArr=NewsListStrArray
		For j =0 to UBound(NewsListStrArray)
			NewsListStrArray(j)=TempArr(UBound(NewsListStrArray)-j)
		Next 
		If Num>0 Then
			TempArr=NewsListStrArray
			For j =0 to Num-1 'UBound(NewsListStrArray)
				NewsListStrArray(j)=TempArr(UBound(NewsListStrArray)-Num+j+1)
			Next 	
		End If 
	End If

	For i = CollectStartLocation to CollectStartLocation + CollectMaxOfOnePage - 1
		if i > UBound(NewsListStrArray) Or (i >= Num And Num<>0) then
			CollectPageNumber = CollectPageNumber + 1
			CollectStartLocation = 0
			CollectedPageURL = ObjURL

			Exit Function
		end If

		AllNewsNumber = AllNewsNumber + 1
		if NewsListStrArray(i) <> "" then
			If BLinkHeadSetting=True Then
				TempArray = GetOtherContent(LinkHeadSetting&NewsListStrArray(i),LinkHeadSetting,LinkFootSetting) 
			ElseIf BLinkFootSetting=True Then 
				TempArray = GetOtherContent(NewsListStrArray(i)&LinkFootSetting,LinkHeadSetting,LinkFootSetting) 
			End If 
			if TempArray <> "" Then
				NewsLinkStr = LoseHtml(FormatUrl(TempArray,CollectObjURL))
				NewsPageStr = GetPageContent(NewsLinkStr)
				if NewsPageStr <> False then		
					TitleStr = LoseHtml(GetOtherContent(NewsPageStr,PageTitleHeadSetting,PageTitleFootSetting))
					Set RsCheckNewsObj = CollectConn.Execute("Select * from FS_News where Links='" & NewsLinkStr & "'")
					if Not RsCheckNewsObj.Eof then
						ReturnValue = GetOneNewsReturnValue(1,i + 1,TitleStr,"",NewsLinkStr) & ReturnValue
					else
						ContentStr = ReplaceKeyWords(GetOneNewsContent(NewsPageStr,NewsLinkStr))
						ContentStr = ReplaceContentStr(ContentStr)
						ContentStr = ReplaceIMGRemoteUrl(ContentStr,SaveIMGPath,AvailableDoMain,DummyDir,NewsLinkStr,SaveRemotePic)
						if TitleStr = "" then
							ReturnValue = GetOneNewsReturnValue(2,i + 1,"","",NewsLinkStr) & ReturnValue
						elseif ContentStr = "" then
							ReturnValue = GetOneNewsReturnValue(3,i + 1,TitleStr,"",NewsLinkStr) & ReturnValue
						else
							ReturnValue = GetOneNewsReturnValue(4,i + 1,TitleStr,ContentStr,NewsLinkStr) & ReturnValue
							if IsNull(HandSetAuthor) OR (HandSetAuthor = "") then
								AuthorStr = LoseHtml(GetOtherContent(NewsPageStr,AuthorHeadSetting,AuthorFootSetting))
							else
								AuthorStr = HandSetAuthor
							end if
							if IsNull(HandSetSource) OR (HandSetSource = "") then
								SourceStr = LoseHtml(GetOtherContent(NewsPageStr,SourceHeadSetting,SourceFootSetting))
							else
								SourceStr = HandSetSource
							end if
							if IsNull(HandSetAddDate) OR Not IsDate(HandSetSource) then
								AddDate = LoseHtml(GetOtherContent(NewsPageStr,AddDateHeadSetting,AddDateFootSetting))
							else
								AddDate = HandSetSource
							end if
							if AddDate <> "" then
								if Not IsDate(AddDate) then	AddDate = Now
							else
								AddDate = Now
							end if
							SaveCollectContent TitleStr,NewsLinkStr,ContentStr,SysClassID,AuthorStr,SourceStr,AddDate
						end if
					end if
					Set RsCheckNewsObj = Nothing
				else
					ReturnValue = GetOneNewsReturnValue(5,i + 1,"","",NewsLinkStr) & ReturnValue
				end if
			else
				ReturnValue = GetOneNewsReturnValue(5,i + 1,"","",NewsLinkStr) & ReturnValue
			end if
		else
			ReturnValue = GetOneNewsReturnValue(5,i + 1,"","",NewsLinkStr) & ReturnValue
		end if
	Next
	CollectStartLocation = i
End Function

Function ResumeGetNewsPageContent()
	dim ResumeSql,RsResumeNewsObj,ResumeNewsURL,ResumeNewsURL1,ResumeNewsLocation
	ResumeSql = "Select top 1 Links from FS_News where SiteID='" & CollectingSiteID &"' order by ID DESC"
	Set RsResumeNewsObj = CollectConn.Execute(ResumeSql)	
	If RsResumeNewsObj.EOF Then 
		set RsResumeNewsObj = nothing
		response.Write("<script>alert(""无法确定您以前的采集记录，\n续采失败！"");history.go(-2);</script>")	
	else
		ResumeNewsURL = RsResumeNewsObj("Links")
		set RsResumeNewsObj = nothing
	End If
	

	Dim NewsPageStr,TitleStr,ContentStr,AuthorStr,SourceStr,AddDate,i,n
	Dim ResponseAllStr,NewsListStr,NewsLinkStr,RsCheckNewsObj
	Dim NewsListStrArray,TempArray
	ResponseAllStr = GetPageContent(FormatUrl(ObjURL,CollectObjURL))	
	if ResponseAllStr = False then
		CollectPageNumber = CollectPageNumber + 1
		ReturnValue = ReturnValue & "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong>错误</strong>:读取新闻列表页面失败<br>"
		Exit Function
	end if

	Dim BLinkHeadSetting,BLinkFootSetting
	BLinkHeadSetting = False
	BLinkFootSetting = False
	If Instr(LinkHeadSetting,"[变量]")<=0 Then
		BLinkHeadSetting = True
	elseif Instr(LinkFootSetting,"[变量]")<=0 Then
		BLinkFootSetting = True
	End If
	If InStr(ResponseAllStr,ListHeadSetting)>0 And InStr(ResponseAllStr,ListFootSetting) Then
		NewsListStr = GetOtherContent(ResponseAllStr,ListHeadSetting,ListFootSetting)
	Else 
		NewsListStr = ResponseAllStr
	End If

	If BLinkHeadSetting Then
		NewsListStr = Mid(NewsListStr,Instr(NewsListStr,LinkHeadSetting)+len(LinkHeadSetting))
		NewsListStrArray = Split(NewsListStr,LinkHeadSetting)
	elseif BLinkFootSetting Then 
		NewsListStr = Left(NewsListStr,InstrRev(NewsListStr,LinkFootSetting))
		NewsListStrArray = Split(NewsListStr,LinkFootSetting)
	End If
	
	For n = 0 to UBound(NewsListStrArray)					
		Dim tempURL
		tempURL=LoseHtml(FormatUrl(GetOtherContent(LinkHeadSetting&NewsListStrArray(n),LinkHeadSetting,LinkFootSetting),CollectObjURL))
		If ResumeNewsURL = tempURL Then
			Exit For
		ElseIf n>=UBound(NewsListStrArray) Then
			AllNewsNumber = AllNewsNumber+n
			CollectPageNumber = CollectPageNumber + 1
			CollectStartLocation = 0
			CollectedPageURL = ObjURL
			Exit Function 			
		End If
	Next 
	CollectStartLocation = n+1

	If IsReverse="1" then '倒序采集
		Dim TempArr,j
		TempArr=NewsListStrArray
		For j =0 to UBound(NewsListStrArray)
			NewsListStrArray(j)=TempArr(UBound(NewsListStrArray)-j)
		Next 
	End If

	For i = CollectStartLocation to CollectStartLocation + CollectMaxOfOnePage - 1
		if i > UBound(NewsListStrArray) Then
			CollectPageNumber = CollectPageNumber + 1
			CollectStartLocation = 0
			CollectedPageURL = ObjURL
			Exit Function
		end If

		AllNewsNumber = AllNewsNumber + 1
		If BLinkHeadSetting Then
			TempArray = GetOtherContent(LinkHeadSetting&NewsListStrArray(i),LinkHeadSetting,LinkFootSetting) 
		elseif BLinkFootSetting Then 
			TempArray = GetOtherContent(NewsListStrArray(i)&LinkFootSetting,LinkHeadSetting,LinkFootSetting) 
		End If  
		if TempArray <> "" Then
			NewsLinkStr = LoseHtml(FormatUrl(TempArray,CollectObjURL))
			Set RsCheckNewsObj = CollectConn.Execute("Select * from FS_News where Links='" & NewsLinkStr & "'")
			if RsCheckNewsObj.Eof then
				NewsPageStr = GetPageContent(NewsLinkStr)
				if NewsPageStr <> False then
					TitleStr = LoseHtml(GetOtherContent(NewsPageStr,PageTitleHeadSetting,PageTitleFootSetting))
				Set RsCheckNewsObj = CollectConn.Execute("Select * from FS_News where Links='" & NewsLinkStr & "'")
					ContentStr = ReplaceKeyWords(GetOneNewsContent(NewsPageStr,NewsLinkStr))
					ContentStr = ReplaceContentStr(ContentStr)
					'不选择远程存图也替换图片地址为绝对地址2005.10.20,通过修改ReplaceIMGRemoteUrl函数
					ContentStr = ReplaceIMGRemoteUrl(ContentStr,SaveIMGPath,AvailableDoMain,DummyDir,NewsLinkStr,SaveRemotePic)
					if TitleStr = "" then
						ReturnValue = GetOneNewsReturnValue(2,i + 1,"","",NewsLinkStr) & ReturnValue
					elseif ContentStr = "" then
						ReturnValue = GetOneNewsReturnValue(3,i + 1,TitleStr,"",NewsLinkStr) & ReturnValue
					else
						ReturnValue = GetOneNewsReturnValue(4,i + 1,TitleStr,ContentStr,NewsLinkStr) & ReturnValue
						if IsNull(HandSetAuthor) OR (HandSetAuthor = "") then
							AuthorStr = LoseHtml(GetOtherContent(NewsPageStr,AuthorHeadSetting,AuthorFootSetting))
						else
							AuthorStr = HandSetAuthor
						end if
						if IsNull(HandSetSource) OR (HandSetSource = "") then
							SourceStr = LoseHtml(GetOtherContent(NewsPageStr,SourceHeadSetting,SourceFootSetting))
						else
							SourceStr = HandSetSource
						end if
						if IsNull(HandSetAddDate) OR Not IsDate(HandSetSource) then
							AddDate = LoseHtml(GetOtherContent(NewsPageStr,AddDateHeadSetting,AddDateFootSetting))
						else
							AddDate = HandSetSource
						end if
						if AddDate <> "" then
							if Not IsDate(AddDate) then	AddDate = Now
						else
							AddDate = Now
						end if
						SaveCollectContent TitleStr,NewsLinkStr,ContentStr,SysClassID,AuthorStr,SourceStr,AddDate
					end if
					Set RsCheckNewsObj = Nothing
				else
					ReturnValue = GetOneNewsReturnValue(5,i + 1,"","",NewsLinkStr) & ReturnValue
				End If
			ElseIf session("ConfirmCollectRevert")<>"ConfirmCollectRevert" Then
				session("ConfirmCollectRevert") = "ConfirmCollectRevert"
				response.write("<script>if(confirm(""您改变过采集顺序吗？\n如果修改过，请单击确定改回原样再续采！\n没有修改过请单击取消继续！""))window.location=""site.asp""</script>")
			End If
		End If		
	Next
	CollectStartLocation = i
End Function

Function GetOneNewsContent(FirstPageContent,NewsLinkStr)
	'On Error Resume Next
	Dim OtherPageNewsLink,OtherPageNewsContentStr,tempSplitArr1,tempSplitArr2
	OtherPageNewsContentStr = FirstPageContent
	GetOneNewsContent = GetOtherContent(FirstPageContent,PagebodyHeadSetting,PagebodyFootSetting)
	if IsNull(OtherNewsPageHeadSetting) OR IsNull(OtherNewsPageFootSetting) OR (OtherNewsPageHeadSetting = "") OR (OtherNewsPageFootSetting = "") Then
		OtherPageNewsLink = ""
	ElseIf  InStr(OtherPageNewsContentStr,OtherNewsPageFootSetting)>0 And InStr(OtherPageNewsContentStr,OtherNewsPageHeadSetting)>0 Then
		tempSplitArr1 = Split(OtherPageNewsContentStr,OtherNewsPageHeadSetting)
		tempSplitArr2 = Split(tempSplitArr1(1),OtherNewsPageFootSetting)
		OtherPageNewsLink = tempSplitArr2(0)
	Else
		OtherPageNewsLink =  GetOtherContent(OtherPageNewsContentStr,OtherNewsPageHeadSetting,OtherNewsPageFootSetting)
	End If
	
	Do While (OtherPageNewsLink <> "")
		OtherPageNewsLink = FormatUrl(OtherPageNewsLink,NewsLinkStr)
		OtherPageNewsContentStr = GetPageContent(OtherPageNewsLink)
		If  InStr(OtherPageNewsContentStr,OtherNewsPageFootSetting)>0 And InStr(OtherPageNewsContentStr,OtherNewsPageHeadSetting)>0 Then
			tempSplitArr1 = Split(OtherPageNewsContentStr,OtherNewsPageHeadSetting)
			tempSplitArr2 = Split(tempSplitArr1(1),OtherNewsPageFootSetting)
			OtherPageNewsLink = tempSplitArr2(0)
		Else
			OtherPageNewsLink =  GetOtherContent(OtherPageNewsContentStr,OtherNewsPageHeadSetting,OtherNewsPageFootSetting)
		End If
		If OtherPageNewsContentStr<>False Then
			GetOneNewsContent = GetOneNewsContent & "[Page]" & GetOtherContent(OtherPageNewsContentStr,PagebodyHeadSetting,PagebodyFootSetting)
		Else
			OtherPageNewsLink = ""
		End If
		If Err Then
			Err.clear
			OtherPageNewsLink = ""
		End If
	Loop
End Function 

Function GetOneNewsContent11(FirstPageContent)'××××××××××××××××××××
	Dim NewsOtherPageContentStr,i
	Dim OtherPageNewsLink,OtherPageNewsContentStr
	if Not IsNull(OtherNewsPageHeadSetting) then
		if OtherNewsPageHeadSetting <> "" then
			if IsNull(OtherNewsPageFootSetting) OR (OtherNewsPageFootSetting = "")Then
				Dim SpaceLoc,BraceLoc
				Dim OtherPageNewsListArray,OtherPageNewsListObjURL
				OtherPageNewsListArray = Split(OtherNewsPageHeadSetting,"</a>")
				For i = LBound(OtherPageNewsListArray) to UBound(OtherPageNewsListArray)
					OtherPageNewsListObjURL = OtherPageNewsListArray(i)
					OtherPageNewsListObjURL = Mid(OtherPageNewsListObjURL,InStr(OtherPageNewsListObjURL,"href") + 4)
					SpaceLoc = InStr(OtherPageNewsListObjURL," ")
					BraceLoc = InStr(OtherPageNewsListObjURL,">")
					if (SpaceLoc <> 0) OR (BraceLoc <> 0) then
						if (SpaceLoc <> 0) And (BraceLoc = 0) then
							OtherPageNewsListObjURL = Left(OtherPageNewsListObjURL,InStr(OtherPageNewsListObjURL," "))
						elseif (SpaceLoc = 0) And (BraceLoc <> 0) then
							OtherPageNewsListObjURL = Left(OtherPageNewsListObjURL,InStr(OtherPageNewsListObjURL,">") - 1)
						else
							if SpaceLoc > BraceLoc then
								OtherPageNewsListObjURL = Left(OtherPageNewsListObjURL,InStr(OtherPageNewsListObjURL,">") - 1)
							else
								OtherPageNewsListObjURL = Left(OtherPageNewsListObjURL,InStr(OtherPageNewsListObjURL," "))
							end if
						end if
					end if
					if OtherPageNewsListObjURL <> "" then
						OtherPageNewsListObjURL = Replace( Replace(Replace(OtherPageNewsListObjURL,"""","")," ",""),"=","")
						OtherPageNewsContentStr = GetPageContent(FormatUrl(OtherPageNewsListObjURL,CollectObjURL))
						if OtherPageNewsContentStr <> False then
							if GetOneNewsContent = "" then
								GetOneNewsContent = GetOtherContent(OtherPageNewsContentStr,PagebodyHeadSetting,PagebodyFootSetting)
							else
								GetOneNewsContent = GetOneNewsContent & "[Page]" & GetOtherContent(OtherPageNewsContentStr,PagebodyHeadSetting,PagebodyFootSetting)
							end if
						end if
					end if
				Next
				if GetOneNewsContent = "" then
					GetOneNewsContent = GetOtherContent(FirstPageContent,PagebodyHeadSetting,PagebodyFootSetting)
				else
					GetOneNewsContent = GetOtherContent(FirstPageContent,PagebodyHeadSetting,PagebodyFootSetting) & "[Page]" & GetOneNewsContent
				end if
			else
				OtherPageNewsLink = GetOtherContent(FirstPageContent,OtherNewsPageHeadSetting,OtherNewsPageFootSetting)
				do while (OtherPageNewsLink <> "" )
					OtherPageNewsContentStr = GetPageContent(FormatUrl(OtherPageNewsLink,CollectObjURL))
					if OtherPageNewsContentStr <> False And InStr(OtherPageNewsContentStr,OtherNewsPageHeadSetting)>0 then
						if GetOneNewsContent = "" then
							GetOneNewsContent = GetOtherContent(OtherPageNewsContentStr,PagebodyHeadSetting,PagebodyFootSetting)
						else
							GetOneNewsContent = GetOneNewsContent & "[Page]" & GetOtherContent(OtherPageNewsContentStr,PagebodyHeadSetting,PagebodyFootSetting)
						end if
						OtherPageNewsLink = GetOtherContent(OtherPageNewsContentStr,OtherNewsPageHeadSetting,OtherNewsPageFootSetting)
					else
						OtherPageNewsLink = ""
					end If
				Loop
				if GetOneNewsContent = "" then
					GetOneNewsContent = GetOtherContent(FirstPageContent,PagebodyHeadSetting,PagebodyFootSetting)
				else
					GetOneNewsContent = GetOtherContent(FirstPageContent,PagebodyHeadSetting,PagebodyFootSetting) & "[Page]" & GetOneNewsContent
				end if
			end if
		else
			GetOneNewsContent = GetOtherContent(FirstPageContent,PagebodyHeadSetting,PagebodyFootSetting)
		end if
	else
		GetOneNewsContent = GetOtherContent(FirstPageContent,PagebodyHeadSetting,PagebodyFootSetting)
	end if
End Function

Function GetOneNewsReturnValue(CauseIndex,NewsIndex,Title,Content,LinkStr)
	Select Case CauseIndex
		Case 1  '不允许重名保存
			GetOneNewsReturnValue = "</br>&nbsp;&nbsp;&nbsp;&nbsp;<strong>序号</strong>： " & NewsIndex
			GetOneNewsReturnValue = GetOneNewsReturnValue & "&nbsp;&nbsp;&nbsp;&nbsp;<strong>结果</strong>： <font color=red>已经采集，在等待审核或者在历史纪录里面</font>"
			GetOneNewsReturnValue = GetOneNewsReturnValue & "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong>标题</strong>： " & Title
			GetOneNewsReturnValue = GetOneNewsReturnValue & "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong>新闻链接</strong>： <a target=""_blank"" href=""" & LinkStr & """>" & LinkStr & "</a><br>"
		Case 2 '标题为空，没有保存
			GetOneNewsReturnValue = "</br>&nbsp;&nbsp;&nbsp;&nbsp;<strong>序号</strong>： " & NewsIndex
			GetOneNewsReturnValue = GetOneNewsReturnValue & "&nbsp;&nbsp;&nbsp;&nbsp;<strong>结果</strong>： <font color=red>标题为空，没有保存</font>"
			GetOneNewsReturnValue = GetOneNewsReturnValue & "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong>新闻链接</strong>： <a target=""_blank"" href=""" & LinkStr & """>" & LinkStr & "</a><br>"
		Case 3 '内容为空，没有保存
			GetOneNewsReturnValue = "</br>&nbsp;&nbsp;&nbsp;&nbsp;<strong>序号</strong>： " & NewsIndex
			GetOneNewsReturnValue = GetOneNewsReturnValue & "&nbsp;&nbsp;&nbsp;&nbsp;<strong>结果</strong>： <font color=red>内容为空，没有保存</font>"
			GetOneNewsReturnValue = GetOneNewsReturnValue & "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong>标题</strong>： " & Title
			GetOneNewsReturnValue = GetOneNewsReturnValue & "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong>新闻链接</strong>： <a target=""_blank"" href=""" & LinkStr & """>" & LinkStr & "</a><br>"
		Case 4 '成功保存
			GetOneNewsReturnValue = "</br>&nbsp;&nbsp;&nbsp;&nbsp;<strong>序号</strong>： " & NewsIndex
			GetOneNewsReturnValue = GetOneNewsReturnValue & "&nbsp;&nbsp;&nbsp;&nbsp;<strong>结果</strong>： 采集成功"
			GetOneNewsReturnValue = GetOneNewsReturnValue & "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong>标题</strong>： " & Title
			GetOneNewsReturnValue = GetOneNewsReturnValue & "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong>内容</strong>： " & Left(LoseHtml(Content),30) & "&nbsp;&nbsp;......"
			GetOneNewsReturnValue = GetOneNewsReturnValue & "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong>新闻链接</strong>： <a target=""_blank"" href=""" & LinkStr & """>" & LinkStr & "</a><br>"
			CollectOKNumber = CollectOKNumber + 1
		Case 5 '不能够读取新闻目标页面
			GetOneNewsReturnValue = "</br>&nbsp;&nbsp;&nbsp;&nbsp;<strong>序号</strong>： " & NewsIndex
			GetOneNewsReturnValue = GetOneNewsReturnValue & "&nbsp;&nbsp;&nbsp;&nbsp;<strong>结果</strong>： <font color=red>不能够读取新闻目标页面</font>"
			GetOneNewsReturnValue = GetOneNewsReturnValue & "<br>&nbsp;&nbsp;&nbsp;&nbsp;<strong>新闻链接</strong>： <a target=""_blank"" href=""" & LinkStr & """>" & LinkStr & "</a><br>"
		Case else
	End Select
End Function

Function SaveCollectContent(Title,Links,Content,ClassID,Author,SourceString,AddDate)
	Dim RsNewsObj,RsTempObj
	Set RsNewsObj = Server.CreateObject("Adodb.RecordSet")
	RsNewsObj.Open "Select * from FS_News where 1=0",CollectConn,3,3
	RsNewsObj.AddNew
	RsNewsObj("Title") = LoseHtml(Title)
	RsNewsObj("Links") = Links
	RsNewsObj("Content") = Content
	RsNewsObj("ContentLength") = Len(Content)
	RsNewsObj("AddDate") = AddDate
	RsNewsObj("ImagesCount") = 0
	RsNewsObj("ClassID") = ClassID
	RsNewsObj("SysTemplet") = SysTemplet
	RsNewsObj("SiteID") = CollectingSiteID
	RsNewsObj("Author") = Left(Author,200)
	RsNewsObj("IsLock") = 0
	RsNewsObj("History") = 0
	RsNewsObj("Source") = Left(SourceString,200)
	RsNewsObj.UpDate
	RsNewsObj.Close
	Set RsNewsObj = Nothing
End Function

Function ReplaceKeyWords(Content)
	Dim RsRuleObj,HeadSeting,FootSeting,ReContent,regEx
	Set RsRuleObj = CollectConn.Execute("Select * from FS_Rule where SiteID=" & CollectingSiteID)
	do while Not RsRuleObj.Eof
		HeadSeting = RsRuleObj("HeadSeting")
		FootSeting = RsRuleObj("FootSeting")
		ReContent = RsRuleObj("ReContent")
		if IsNull(FootSeting) or FootSeting = "" then
			if HeadSeting <> "" then
				Content = Replace(Content,HeadSeting,ReContent)
			end if
		end if
		if Not IsNull(FootSeting) and FootSeting <> "" and Not IsNull(HeadSeting) and HeadSeting <> ""  then
			Set regEx = New RegExp
			regEx.Pattern = HeadSeting & "[^\0]*" & FootSeting
			regEx.IgnoreCase = False
			regEx.Global = True
			'Dim Matches,Match,HaveTF,ShowStr
			'HaveTF = False
			'Set Matches = regEx.Execute(Content)
				'For Each Match in Matches
					'ShowStr = ShowStr & Match.Value & "<br>"
					'HaveTF = True
				'Next
			'if HaveTF = True then
				'Response.Write(ShowStr)
				'Response.End
			'end if
			if IsNull(ReContent) then
				Content = regEx.Replace(Content,"")
			else
				Content = regEx.Replace(Content,ReContent)
			end if
			Set regEx = Nothing
		end if
		RsRuleObj.MoveNext
	loop
	Set RsRuleObj = Nothing
	ReplaceKeyWords = Content
End Function
%>