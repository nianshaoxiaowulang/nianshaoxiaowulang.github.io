<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
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
'判断权限
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P080100") then Call ReturnError1()
'判断权限结束
Dim SelectPath
if SysRootDir = "" then
	SelectPath = "/" & UpFiles
else
	SelectPath = "/" & SysRootDir & "/" & UpFiles
end if
Dim Rs
if Request("Action") = "Del" then
	if Not JudgePopedomTF(Session("Name"),"P080103") then Call ReturnError1()
	if Request("id") <> "" then
		CollectConn.Execute("delete from FS_Site where ID in (" & Replace(Request("id"),"***",",") & ")")
	end If
	if Request("SiteFolderID") <> "" then
		CollectConn.Execute("delete from FS_SiteFolder where ID in (" & Replace(Request("SiteFolderID"),"***",",") & ")")
	end if
	Response.Redirect("site.asp")
	Response.End
elseif Request("Action") = "Lock" then
	if Request("LockID") <> "" then
		CollectConn.Execute("Update FS_Site Set IsLock=1 where ID in (" & Replace(Request("LockID"),"***",",") & ")")
		Response.Redirect("site.asp")
		Response.End
	end if
elseif Request("Action") = "UNLock" then
	if Request("LockID") <> "" then
		CollectConn.Execute("Update FS_Site Set IsLock=0 where ID in (" & Replace(Request("LockID"),"***",",") & ")")
		Response.Redirect("site.asp")
		Response.End
	end if
end if
if Request.Form("vs")="add" then
	if Not JudgePopedomTF(Session("Name"),"P080101") then Call ReturnError1()
    if Request.Form("SaveIMGPath") = "" OR Request.Form("SiteName")="" Or Request.Form("SysTemplet")="" or Request.Form("objURL")=""  or Request.Form("SysClass")=""  then
		Response.write"<script>alert(""请填写完整！"");location.href=""javascript:history.back()"";</script>"
		Response.end
	end if
    Dim Sql
	Set Rs = Server.CreateObject ("ADODB.RecordSet")
	Sql = "Select * from FS_Site where 1=0"
	Rs.Open Sql,CollectConn,1,3
	Rs.AddNew
	Rs("SiteName") = NoCSSHackAdmin(Request.Form("SiteName"),"站点名称")
	Rs("SysTemplet") = Request.Form("SysTemplet")
	Rs("objURL") = Request.Form("objURL")
	Rs("folder") = Request.Form("SiteFolder")
	Rs("SysClass") = Request.Form("SysClass")
	Rs("SaveIMGPath") = Request.Form("SaveIMGPath")
	if Request.Form("IsIFrame") = "1" then
		Rs("IsIFrame") = True
	else
		Rs("IsIFrame") = False
	end if
	if Request.Form("IsReverse") = "1" then
		Rs("IsReverse") = 1
	else
		Rs("IsReverse") = 0
	end if
	if Request.Form("IsScript") = "1" then
		Rs("IsScript") = True
	else
		Rs("IsScript") = False
	end if
	if Request.Form("IsClass") = "1" then
		Rs("IsClass") = True
	else
		Rs("IsClass") = False
	end if
	if Request.Form("IsFont") = "1" then
		Rs("IsFont") = True
	else
		Rs("IsFont") = False
	end if
	if Request.Form("IsSpan") = "1" then
		Rs("IsSpan") = True
	else
		Rs("IsSpan") = False
	end if
	if Request.Form("IsObject") = "1" then
		Rs("IsObject") = True
	else
		Rs("IsObject") = False
	end if
	if Request.Form("IsStyle") = "1" then
		Rs("IsStyle") = True
	else
		Rs("IsStyle") = False
	end if
	if Request.Form("IsDiv") = "1" then
		Rs("IsDiv") = True
	else
		Rs("IsDiv") = False
	end if
	if Request.Form("IsA") = "1" then
		Rs("IsA") = True
	else
		Rs("IsA") = False
	end if
	if Request.Form("Audit") = "1" then
		Rs("Audit") = True
	else
		Rs("Audit") = False
	end if
	if Request.Form("TextTF") = "1" then
		Rs("TextTF") = True
	else
		Rs("TextTF") = False
	end if
	if Request.Form("SaveRemotePic") = "1" then
		Rs("SaveRemotePic") = True
	else
		Rs("SaveRemotePic") = False
	end if
	if Request.Form("Islock") <> "" then
		Rs("Islock") = True
	else
		Rs("Islock") = False
	end if
	Rs.UpDate
	Rs.Close
	Set Rs = Nothing
	Set Conn = Nothing
	Set CollectConn = Nothing
	Response.Redirect("Site.asp")
	Response.End
elseif Request("vs")="addfolder" then
	if Not JudgePopedomTF(Session("Name"),"P080101") then Call ReturnError1()
	Dim SiteFolder,SiteFolderDetail,SqlStr
	SiteFolder = NoCSSHackAdmin(Request.Form("SiteFolder"),"站点栏目")
	SiteFolderDetail = Request.Form("SiteFolderDetail")
	If SiteFolder = "" or SiteFolderDetail = "" Then
		Response.write"<script>alert(""请填写完整！"");location.href=""javascript:history.back()"";</script>"
		Response.end
	End If
	Set Rs = Server.CreateObject ("ADODB.RecordSet")
	SqlStr = "Select * from FS_SiteFolder where 1=0"
	Rs.Open SqlStr,CollectConn,1,3
	Rs.AddNew
	Rs("SiteFolder") = SiteFolder
	Rs("SiteFolderDetail") = SiteFolderDetail
	Rs.UpDate
	Rs.Close
	Set Rs = Nothing
	Set Conn = Nothing
	Set CollectConn = Nothing
	Response.Redirect("Site.asp")
	Response.end
elseif Request("vs")="Copy" then
	if Not JudgePopedomTF(Session("Name"),"P080104") then Call ReturnError1()
	Dim SiteID,SiteFolderID,RsCopySourceObj,RsCopyObjectObj,FiledObj
	SiteID = Request("SiteID")
	SiteFolderID = Request("SiteFolderID")
	if SiteID <> "" then
		Set RsCopySourceObj = CollectConn.Execute("Select * from FS_Site where ID in (" & Replace(SiteID,"***",",") & ")")
		do while Not RsCopySourceObj.Eof
			Set RsCopyObjectObj = Server.CreateObject("ADODB.RecordSet")
			RsCopyObjectObj.Open "Select * from FS_Site where 1=0",CollectConn,3,3
			RsCopyObjectObj.AddNew
			For Each FiledObj In RsCopyObjectObj.Fields
				if LCase(FiledObj.name) <> "id" then
					RsCopyObjectObj(FiledObj.name) = RsCopySourceObj(FiledObj.name)
				end if
			Next
			RsCopyObjectObj.Update
			RsCopySourceObj.MoveNext
		Loop
		Set RsCopySourceObj = Nothing
		Set RsCopyObjectObj = Nothing
	end If
	if SiteFolderID <> "" then
		Set RsCopySourceObj = CollectConn.Execute("Select * from FS_SiteFolder where ID in (" & Replace(SiteFolderID,"***",",") & ")")
		do while Not RsCopySourceObj.Eof
			Set RsCopyObjectObj = Server.CreateObject("ADODB.RecordSet")
			RsCopyObjectObj.Open "Select * from FS_SiteFolder where 1=0",CollectConn,3,3
			RsCopyObjectObj.AddNew
			For Each FiledObj In RsCopyObjectObj.Fields
				if LCase(FiledObj.name) <> "id" then
					RsCopyObjectObj(FiledObj.name) = RsCopySourceObj(FiledObj.name)
				end if
			Next
			RsCopyObjectObj.Update
			RsCopySourceObj.MoveNext
		Loop
		Set RsCopySourceObj = Nothing
		Set RsCopyObjectObj = Nothing
	end if
	Set Conn = Nothing
	Set CollectConn = Nothing
	Response.Redirect("Site.asp")
	Response.end
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>自动新闻采集—站点设置</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<% if Request("Action") <> "Addsite" and Request("Action") <> "Addsitefolder" then %>
<script language="JavaScript" src="../../SysJS/ContentMenu.js"></script>
<% end if %>
<body<% if Request("Action") <> "Addsite" and Request("Action") <> "Addsitefolder" then %> onselectstart="return false;" onClick="SelectSite();"<% end if %> leftmargin="2" topmargin="2">
<%
if Request("Action") = "Addsite" then
	Call Add()
ElseIf Request("Action") = "Addsitefolder" Then
	Call AddFolder()
ElseIf Request("Action") = "SubFolder" Then
	Call SubMain()
Else
	Call Main()
end if
Sub Main()
	Session("SessionReturnValue") = ""
	if Not JudgePopedomTF(Session("Name"),"P080100") then Call ReturnError1()
%>
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=55 align="center" alt="添加采集栏目" onClick="AddSiteFolder();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">新建栏目</td>
		  <td width=2 class="Gray">|</td>
		  <td width=55 align="center" alt="添加采集站点" onClick="AddSite();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">新建站点</td>
		  <td width=2 class="Gray">|</td>
		  <td width=55 align="center" alt="修改站点属性" onClick="EditSite();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">修改属性</td>
		  <td width=2 class="Gray">|</td>
		  <td width=55 align="center" alt="修改站点向导" onClick="EditSiteGuide();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">修改向导</td>
		  <td width=2 class="Gray">|</td>
          <td width=35 align="center" alt="删除站点" onClick="DelSite();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">删除</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="复制站点" onClick="CopySite();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">复制</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="开始采集" onClick="StartCollect();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">采集</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="继续上次采集" onClick="ResumeCollect();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">续采</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
		  <td>&nbsp; </td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
  	<td colspan="5" height="2"></td>
  </tr>
  <tr> 
    <td width="19%" height="26" nowrap bgcolor="#FFFFFF" class="ButtonListLeft"> <div align="center">名称</div></td>
    <td width="9%" height="26" nowrap bgcolor="#FFFFFF" class="ButtonList"> <div align="center">状态</div></td>
    <td width="9%" height="26" bgcolor="#FFFFFF" class="ButtonList" nowrap> <div align="center">采集对象页</div></td>
    <td width="20%" height="26" nowrap bgcolor="#FFFFFF" class="ButtonList"> <div align="center">采集到栏目</div></td>
    <td width="12%" height="26" nowrap class="ButtonList"> <div align="center">开始采集</div></td>
  </tr>
  <%
	Dim RsSite,SiteSql,CheckInfo
	Dim RsSiteFolder
	Set RsSiteFolder = CollectConn.Execute("select * from FS_SiteFolder order by id DESC")
	Do While not RsSiteFolder.EOF
	%>
  <tr title="站点分类管理目录，点击进入"> 
    <td height="26" nowrap> <table border="0" cellspacing="0" cellpadding="0">
        <tr> 	
          <td><img src="../../Images/Folder/folderclosed.gif" width="24" height="22"></td>
          <td nowrap><span  class="TempletItem" SiteFolderID="<% = RsSiteFolder("ID") %>" onDblClick="ChangeFolder(this);"> 
            <%= RsSiteFolder("SiteFolder")%>
            </span></td>
			 </tr>
      </table></td>
    <td nowrap> <div align="center"> &nbsp; </div></td>
    <td nowrap> <div align="center"> &nbsp; </div></td>
    <td nowrap> <div align="center"> &nbsp; </div></td>
    <td nowrap> <div align="center"> &nbsp;  </div></td>
		 </tr>
	<%
		RsSiteFolder.MoveNext
	Loop
	Set RsSiteFolder = Nothing


	Dim IsCollect,SysClassCName,RsTempObj,CollectPromptInfo
	Set RsSite = Server.CreateObject ("ADODB.RecordSet")
	SiteSql="Select * from FS_Site where folder=0 order by id desc"
		RsSite.Open SiteSql,CollectConn,1,1
	Do While not RsSite.eof
		if  RsSite("LinkHeadSetting") <> "" And  RsSite("LinkFootSetting") <> "" And RsSite("PagebodyHeadSetting") <> "" And  RsSite("PagebodyFootSetting") <> "" And  RsSite("PageTitleHeadSetting") <> "" And  RsSite("PageTitleFootSetting") <> "" then
			if RsSite("IsLock") = True then
				IsCollect = False
				CollectPromptInfo = "站点已经被锁定,不能采集"
			else
				IsCollect = True
				CollectPromptInfo = "可以采集,请检查是否设置正确，否则不能进行采集"
			end if
		else
			IsCollect = False
			CollectPromptInfo = "不能采集,请把匹配规则设置完整"
		end if
		
		Set RsTempObj = Conn.Execute("Select ClassCName from FS_NewsClass where ClassID='" & RsSite("SysClass") & "'")
		if Not RsTempObj.Eof then
			SysClassCName = RsTempObj("ClassCName")
		else
			SysClassCName = "栏目不存在"
			IsCollect = False
			CollectPromptInfo = "目标栏目不存在,不能采集"
		end if
		Set RsTempObj = Nothing
	%>
	<tr title="<% = CollectPromptInfo %>"> 
    <td height="26" nowrap> <table border="0" cellspacing="0" cellpadding="0">
        <tr> 
		<td><img src="../../Images/Station.gif" width="24" height="22"></td>
          <td nowrap><span IsCollect="<% = IsCollect %>" class="TempletItem" SiteID="<% = RsSite("ID") %>"> 
            <% = RsSite("SiteName") %>
            </span></td>
		 </tr>
      </table></td>
    <td nowrap> <div align="center"> 
        <%
		if RsSite("IsLock") = True then
			Response.Write("锁定")
		ElseIf IsCollect = False Then
			Response.Write("无效")
		else
			Response.Write("有效")
		end if
		%>
      </div></td>
    <td nowrap> <div align="center"><a href="<% = RsSite("objURL") %>" target="_blank"><img src="Images/objpage.gif" alt="点击访问" width="20" height="20" border="0"></a></div></td>
    <td nowrap> <div align="center"> 
        <% = SysClassCName %>
      </div></td>
    <td nowrap> <div align="center"> 
        <% if IsCollect = true then %>
        <div align="center"><span onClick="StartOneSiteCollect('<% = RsSite("Id") %>');" style="cursor:hand;"><img src="Images/collect.gif" width="20" height="20" border="0"></span></div>
        <% else %>
        <div align="center"><img src="Images/uncollect.gif" width="20" height="20" border="0"></div>
        <% end if %>
      </div></td>
  </tr>
  <%
		RsSite.MoveNext
	loop
	RsSite.close
	Set RsSite = Nothing
end Sub

Sub	SubMain()
	Session("SessionReturnValue") = ""
	if Not JudgePopedomTF(Session("Name"),"P080100") then Call ReturnError1()
%>
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=55 align="center" alt="添加采集栏目" onClick="AddSiteFolder();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">新建栏目</td>
		  <td width=2 class="Gray">|</td>
		  <td width=55 align="center" alt="添加采集站点" onClick="AddSubSite();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">新建站点</td>
		  <td width=2 class="Gray">|</td>
		  <td width=55 align="center" alt="修改站点属性" onClick="EditSite();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">修改属性</td>
		  <td width=2 class="Gray">|</td>
		  <td width=55 align="center" alt="修改站点向导" onClick="EditSiteGuide();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">修改向导</td>
		  <td width=2 class="Gray">|</td>
          <td width=35 align="center" alt="删除站点" onClick="DelSite();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">删除</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="复制站点" onClick="CopySite();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">复制</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="开始采集" onClick="StartCollect();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">采集</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="继续上次采集" onClick="ResumeCollect();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">继采</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
		  <td>&nbsp; </td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <script language="javascript">
	function AddSubSite()
	{
		window.location="site.asp?Action=Addsite&FolderID="+<%=Request("FolderID")%>;
	}
  </script>
  <tr>
  	<td colspan="5" height="2"></td>
  </tr>
  <tr> 
    <td width="19%" height="26" nowrap bgcolor="#FFFFFF" class="ButtonListLeft"> <div align="center">名称</div></td>
    <td width="9%" height="26" nowrap bgcolor="#FFFFFF" class="ButtonList"> <div align="center">状态</div></td>
    <td width="9%" height="26" bgcolor="#FFFFFF" class="ButtonList" nowrap> <div align="center">采集对象页</div></td>
    <td width="20%" height="26" nowrap bgcolor="#FFFFFF" class="ButtonList"> <div align="center">采集到栏目</div></td>
    <td width="12%" height="26" nowrap class="ButtonList"> <div align="center">开始采集</div></td>
  </tr>
  <tr title="点击返回上级目录"> 
    <td height="26" nowrap> <table border="0" cellspacing="0" cellpadding="0">
        <tr> 	
          <td><img src="../../Images/arrow.gif" width="24" height="22"></td>
          <td nowrap><span  class="TempletItem" onDblClick="backTop()" onclick="backTop()" style="cursor:hand">返回上级</span></td>
			 </tr>
      </table></td>
    <td nowrap> <div align="center"> &nbsp; </div></td>
    <td nowrap> <div align="center"> &nbsp; </div></td>
    <td nowrap> <div align="center"> &nbsp; </div></td>
    <td nowrap> <div align="center"> &nbsp;  </div></td>
 </tr>
  <%
	Dim RsSite,SiteSql,CheckInfo
	Dim IsCollect,SysClassCName,RsTempObj,CollectPromptInfo
	Set RsSite = Server.CreateObject ("ADODB.RecordSet")
	SiteSql="Select * from FS_Site where folder="& CLng(Request("FolderID")) &" order by id desc"
		RsSite.Open SiteSql,CollectConn,1,1
	Do While not RsSite.eof
		if  RsSite("LinkHeadSetting") <> "" And  RsSite("LinkFootSetting") <> "" And RsSite("PagebodyHeadSetting") <> "" And  RsSite("PagebodyFootSetting") <> "" And  RsSite("PageTitleHeadSetting") <> "" And  RsSite("PageTitleFootSetting") <> "" then
			if RsSite("IsLock") = True then
				IsCollect = False
				CollectPromptInfo = "站点已经被锁定,不能采集"
			else
				IsCollect = True
				CollectPromptInfo = "可以采集,请检查是否设置正确，否则不能进行采集"
			end if
		else
			IsCollect = False
			CollectPromptInfo = "不能采集,请把匹配规则设置完整"
		end if
		
		Set RsTempObj = Conn.Execute("Select ClassCName from FS_NewsClass where ClassID='" & RsSite("SysClass") & "'")
		if Not RsTempObj.Eof then
			SysClassCName = RsTempObj("ClassCName")
		else
			SysClassCName = "栏目不存在"
			IsCollect = False
			CollectPromptInfo = "目标栏目不存在,不能采集"
		end if
		Set RsTempObj = Nothing
	%>
	<tr title="<% = CollectPromptInfo %>"> 
    <td height="26" nowrap> <table border="0" cellspacing="0" cellpadding="0">
        <tr> 
		<td><img src="../../Images/Station.gif" width="24" height="22"></td>
          <td nowrap><span IsCollect="<% = IsCollect %>" class="TempletItem" SiteID="<% = RsSite("ID") %>"> 
            <% = RsSite("SiteName") %>
            </span></td>
		 </tr>
      </table></td>
    <td nowrap> <div align="center"> 
        <%
		if RsSite("IsLock") = True then
			Response.Write("锁定")
		ElseIf IsCollect = False Then
			Response.Write("无效")
		else
			Response.Write("有效")
		end if
		%>
      </div></td>
    <td nowrap> <div align="center"><a href="<% = RsSite("objURL") %>" target="_blank"><img src="Images/objpage.gif" alt="点击访问" width="20" height="20" border="0"></a></div></td>
    <td nowrap> <div align="center"> 
        <% = SysClassCName %>
      </div></td>
    <td nowrap> <div align="center"> 
        <% if IsCollect = true then %>
        <div align="center"><span onClick="StartOneSiteCollect('<% = RsSite("Id") %>');" style="cursor:hand;"><img src="Images/collect.gif" width="20" height="20" border="0"></span></div>
        <% else %>
        <div align="center"><img src="Images/uncollect.gif" width="20" height="20" border="0"></div>
        <% end if %>
      </div></td>
  </tr>
  <%
		RsSite.MoveNext
	loop
	RsSite.close
	Set RsSite = Nothing
end sub

Sub Add()
	if Not JudgePopedomTF(Session("Name"),"P080101") then Call ReturnError1()
	Dim TempletDirectory
	if SysRootDir <> "" then
		TempletDirectory = "/" & SysRootDir & "/" & TempletDir
	else
		TempletDirectory = "/" & TempletDir
	end if
%>
<form name="AddSiteForm" method="post" action="">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="保存" onClick="document.AddSiteForm.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
            <td>&nbsp; <input name="vs" type="hidden" id="vs2" value="add"> </td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="dddddd">
    <tr bgcolor="#FFFFFF"> 
      <td width="100" height="26"> 
        <div align="center">采集站点名称</div></td>
      <td> 
        <input name="SiteName" style="width:100%;" type="text" id="SiteName2"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26"> 
        <div align="center">采集站点分类</div></td>
      <td> 
        <select name="SiteFolder" style="width:100%;" id="SiteFolder">
		<option value="0">根栏目</option>
          <% = FolderList %>
        </select></td>
    </tr>
	<tr bgcolor="#FFFFFF"> 
      <td height="26"> 
        <div align="center">入库目标栏目</div></td>
      <td> 
        <select name="SysClass" style="width:100%;" id="select">
          <% = ClassList %>
        </select></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26"> 
        <div align="center">采集对象页</div></td>
      <td> 
        <input style="width:100%;" name="objURL" type="text" id="objURL" value="http://"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26"> 
        <div align="center">新闻摸板</div></td>
      <td> 
        <input readonly name="SysTemplet" type="text" id="SysTemplet" style="width:80%;"> 
        <input name="Submitaaa" type="button" id="Submitaaa" value="选择模板" onClick="OpenWindowAndSetValue('../../FunPages/SelectFileFrame.asp?CurrPath=<% = TempletDirectory %>',400,300,window,document.AddSiteForm.SysTemplet);"> 
        <div align="right"></div></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26"> 
        <div align="center">采集参数</div></td>
      <td>锁定 
        <input name="islock" type="checkbox" id="islock" value="1">
        保存远程图片 
        <input type="checkbox" name="SaveRemotePic" value="1">
        新闻是否已经审核 
        <input name="Audit" type="checkbox" value="1" checked>
        是否倒序采集 
        <input name="IsReverse" type="checkbox" id="IsReverse" value="1"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26"> 
        <div align="center">保存图片路径</div></td>
      <td> 
        <input type="text" readonly name="SaveIMGPath" style="width:80%;" value="/<% = UpFiles & "/" & BeyondPicDir %>">
        <input name="Submit111" id="SelectPath" type="button" value="选择路径" onClick="OpenWindowAndSetValue('../../FunPages/SelectPathFrame.asp?CurrPath=<% = SelectPath %>',400,300,window,document.AddSiteForm.SaveIMGPath);"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26"> 
        <div align="center">过滤选项</div></td>
      <td>HTML 
        <input type="checkbox" name="TextTF" value="1">
        STYLE <input type="checkbox" name="IsStyle" value="1">
        DIV<input type="checkbox" name="IsDiv" value="1">
        A<input type="checkbox" name="IsA" value="1">
        CLASS<input type="checkbox" name="IsClass" value="1">
        FONT<input type="checkbox" name="IsFont" value="1">
        SPAN<input type="checkbox" name="IsSpan" value="1">
        OBJECT<input type="checkbox" name="IsObject" value="1">
        IFRAME<input type="checkbox" name="IsIFrame" value="1">
        SCRIPT<input type="checkbox" name="IsScript" value="1">
        </td>
    </tr>
</table>
</form>
<% End Sub

Sub AddFolder()
	if Not JudgePopedomTF(Session("Name"),"P080101") then Call ReturnError1()
	Dim TempletDirectory
	if SysRootDir <> "" then
		TempletDirectory = "/" & SysRootDir & "/" & TempletDir
	else
		TempletDirectory = "/" & TempletDir
	end if
%>
<form name="AddSiteFolderForm" method="post" action="">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="保存" onClick="document.AddSiteFolderForm.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
          <td>&nbsp; <input name="vs" type="hidden" id="vs2" value="addfolder"> </td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="dddddd">
    <%
	Dim SiteFolderID,RsSiteFolder
	SiteFolderID = Request("SiteFolderID")
	If SiteFolderID<>"" Then
		Set RsSiteFolder = CollectConn.Execute("select * from FS_SiteFolder where ID=" & SiteFolderID)
		If RsSiteFolder.EOF Then
			Response.write "<script>alert('该栏目不存在！');history.back();</script>"
			Response.End
		End If
	%>
	<tr bgcolor="#FFFFFF"> 
      <td width="100" height="26"> 
        <div align="center">栏目名称:</div></td>
	  <td> 
        <input style="width:100%" type="text" name="SiteFolder" value="<%=RsSiteFolder("SiteFolder")%>"></td>
	</tr>
	<tr bgcolor="#FFFFFF"> 
      <td width="100" height="26"> 
        <div align="center">栏目说明:</div></td>
	  <td> 
        <textarea style="width:100%" name="SiteFolderDetail" rows="10"><%=RsSiteFolder("SiteFolderDetail")%></textarea></td>
	</tr>
	<%
	Else
	%>
	<tr bgcolor="#FFFFFF"> 
      <td width="100" height="26"> 
        <div align="center">栏目名称:</div></td>
	  <td> 
        <input style="width:100%" type="text" name="SiteFolder"></td>
	</tr>
	<tr bgcolor="#FFFFFF"> 
      <td width="100" height="26"> 
        <div align="center">栏目说明:</div></td>
	  <td> 
        <textarea style="width:100%" name="SiteFolderDetail" rows="10"></textarea></td>
	</tr>
	<%
	End If
	%>
</table>
</form>
<% End Sub %>

</body>
</html>
<%
Function ClassList()
	Dim ClassListObj
	Set ClassListObj = Conn.Execute("Select ClassID,ClassCName from FS_NewsClass where ParentID='0' order by ClassID desc")
	do while Not ClassListObj.Eof
		ClassList = ClassList & "<option value="&ClassListObj("ClassID")&"" & ">" & ClassListObj("ClassCName") & "</option><br>"
		ClassList = ClassList & ChildClassList(ClassListObj("ClassID"),"")
		ClassListObj.MoveNext	
	loop
	ClassListObj.Close
	Set ClassListObj = Nothing
End Function

Function FolderList()
	Dim FolderListObj,StrSelected
	Set FolderListObj = Collectconn.Execute("Select * from FS_SiteFolder order by ID desc")
	do while Not FolderListObj.Eof
		If CInt(Request("FolderID"))=FolderListObj("ID") Then
			StrSelected="selected"
		Else
			StrSelected=""
		End If
		FolderList = FolderList & "<option value="&FolderListObj("ID")&" " & StrSelected & ">&nbsp;&nbsp;|--" & FolderListObj("SiteFolder") & "</option><br>"
		FolderListObj.MoveNext	
	loop
	FolderListObj.Close
	Set FolderListObj = Nothing
End Function

Function ChildClassList(ClassID,Temp)
	Dim TempRs,TempStr
	Set TempRs = Conn.Execute("Select ClassID,ClassCName from FS_NewsClass where ParentID='" & ClassID & "' order by ClassID desc")
	TempStr = Temp & " |- "
	do while Not TempRs.Eof
		ChildClassList = ChildClassList & "<option value="&TempRs("ClassID")&"" & ">" & TempStr & TempRs("ClassCName") & "</option><br>"
		ChildClassList = ChildClassList & ChildClassList(TempRs("ClassID"),TempStr)
		TempRs.MoveNext
	loop
	TempRs.Close
	Set TempRs = Nothing
End Function

Set Conn = Nothing
Set CollectConn = Nothing
%>
<script language="JavaScript">
var DocumentReadyTF=false;
var ListObjArray = new Array();
var ContentMenuArray=new Array();
function document.onreadystatechange()
{
	if (DocumentReadyTF) return;
	<% if Request("Action") <> "Addsite" and Request("Action") <> "AddsiteFolder" then %>
	IntialListObjArray();
	InitialContentListContentMenu();
	<% end if %>
	DocumentReadyTF=true;
}
function InitialContentListContentMenu()
{
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.AddSite();",'新建站点','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.EditSite();",'修改','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.DelSite();",'删除','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.StartCollect();','采集','disabled');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.ResumeCollect();','续采','disabled');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.StartCollectAtServer();','后台采集','disabled');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.EditSiteGuide();','向导','disabled');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.CopySite();','复制','disabled');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.Lock(true);','锁定','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.Lock(false);','解锁','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	//ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.Export();','导出','disabled');
	//ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.Import();','导入','disabled');
	//ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('top.GetEkMainObject().location.reload();','刷新','');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('prompt(\'本页面路径属性\',\'<%=Request.ServerVariables("SCRIPT_NAME")%>\');','路径属性','');
}
function ContentMenuFunction(ExeFunction,Description,EnabledStr)
{
	this.ExeFunction=ExeFunction;
	this.Description=Description;
	this.EnabledStr=EnabledStr;
}
function ContentMenuShowEvent()
{
	ChangeContentMenuStatus();
}
function ChangeContentMenuStatus()
{
	var EventObjInArray=false,SelectContent='',SelectFolder='',DisabledContentMenuStr='';
	for (var i=0;i<ListObjArray.length;i++)
	{
		if (event.srcElement==ListObjArray[i].Obj)
		{
			if (ListObjArray[i].Selected==true) EventObjInArray=true;
			break;
		}
	}
	for (var i=0;i<ListObjArray.length;i++)
	{
		if (event.srcElement==ListObjArray[i].Obj)
		{
			ListObjArray[i].Obj.className='TempletSelectItem';
			ListObjArray[i].Selected=true;
			if (ListObjArray[i].Obj.SiteID!=null)
			{
				if (SelectContent=='') SelectContent=ListObjArray[i].Obj.SiteID;
				else SelectContent=SelectContent+'***'+ListObjArray[i].Obj.SiteID;
			}
			if (ListObjArray[i].Obj.SiteFolderID!=null)
			{
				if (SelectFolder=='') SelectFolder=ListObjArray[i].Obj.SiteFolderID;
				else SelectFolder=SelectFolder+'***'+ListObjArray[i].Obj.SiteFolderID;
			}
		}
		else
		{
			if (!EventObjInArray)
			{
				ListObjArray[i].Obj.className='TempletItem';
				ListObjArray[i].Selected=false;
			}
			else
			{
				if (ListObjArray[i].Selected==true)
				{
					if (ListObjArray[i].Obj.SiteID!=null)
					{
						if (SelectContent=='') SelectContent=ListObjArray[i].Obj.SiteID;
						else SelectContent=SelectContent+'***'+ListObjArray[i].Obj.SiteID;
					}
					if (ListObjArray[i].Obj.SiteFolderID!=null)
					{
						if (SelectFolder=='') SelectFolder=ListObjArray[i].Obj.SiteFolderID;
						else SelectFolder=SelectFolder+'***'+ListObjArray[i].Obj.SiteFolderID;
					}
				}
			}
		}
	}
	if (SelectContent=='' && SelectFolder=='') DisabledContentMenuStr=',修改,删除,向导,复制,采集,续采,后台采集,锁定，解锁,导出,';
	else if (SelectContent!='' && SelectFolder!='')
	{
		DisabledContentMenuStr=',修改,向导,采集,续采,后台采集,锁定，解锁,导出,';
	}
	else if (SelectContent!='')
	{
		if (SelectContent.indexOf('***')==-1) DisabledContentMenuStr='';
		else DisabledContentMenuStr=',修改,向导,复制,';
	}
	else
	{		
		if (SelectFolder.indexOf('***')==-1) DisabledContentMenuStr=',向导,采集,续采,后台采集,锁定，解锁,导出,';
		else DisabledContentMenuStr=',修改,向导,采集,续采,后台采集,锁定，解锁,导出,';
	}
	for (var i=0;i<ContentMenuArray.length;i++)
	{
		if (DisabledContentMenuStr.indexOf(ContentMenuArray[i].Description)!=-1) ContentMenuArray[i].EnabledStr='disabled';
		else  ContentMenuArray[i].EnabledStr='';
	}
}
function FolderFileObj(Obj,Index,Selected)
{
	this.Obj=Obj;
	this.Index=Index;
	this.Selected=Selected;
}
function IntialListObjArray()
{
	var CurrObj=null,j=1;
	for (var i=0;i<document.all.length;i++)
	{
		CurrObj=document.all(i);
		if (CurrObj.SiteID!=null)
		{
			ListObjArray[ListObjArray.length]=new FolderFileObj(CurrObj,j,false);
			j++;
		}
		if (CurrObj.SiteFolderID!=null)
		{
			ListObjArray[ListObjArray.length]=new FolderFileObj(CurrObj,j,false);
			j++;
		}
	}
}
function SelectSite()
{
	var el=event.srcElement;
	var i=0;
	if ((event.ctrlKey==true)||(event.shiftKey==true))
	{
		if (event.ctrlKey==true)
		{
			for (i=0;i<ListObjArray.length;i++)
			{
				if (el==ListObjArray[i].Obj)
				{
					if (ListObjArray[i].Selected==false)
					{
						ListObjArray[i].Obj.className='TempletSelectItem';
						ListObjArray[i].Selected=true;
					}
					else
					{
						ListObjArray[i].Obj.className='TempletItem';
						ListObjArray[i].Selected=false;
					}
				}
			}
		}
		if (event.shiftKey==true)
		{
			var MaxIndex=0,ObjInArray=false,EndIndex=0,ElIndex=-1;
			for (i=0;i<ListObjArray.length;i++)
			{
				if (ListObjArray[i].Selected==true)
				{
					if (ListObjArray[i].Index>=MaxIndex) MaxIndex=ListObjArray[i].Index;
				}
				if (el==ListObjArray[i].Obj)
				{
					ObjInArray=true;
					ElIndex=i;
					EndIndex=ListObjArray[i].Index;
				}
			}
			if (ElIndex>MaxIndex)
				for (i=MaxIndex-1;i<EndIndex;i++)
				{
					ListObjArray[i].Obj.className='TempletSelectItem';
					ListObjArray[i].Selected=true;
				}
			else
			{
				for (i=EndIndex;i<MaxIndex-1;i++)
				{	
					ListObjArray[i].Obj.className='TempletSelectItem';
					ListObjArray[i].Selected=true;
				}
				ListObjArray[ElIndex].Obj.className='TempletSelectItem';
				ListObjArray[ElIndex].Selected=true;
			}
		}
	}
	else
	{
		for (i=0;i<ListObjArray.length;i++)
		{
			if (el==ListObjArray[i].Obj)
			{
				ListObjArray[i].Obj.className='TempletSelectItem';
				ListObjArray[i].Selected=true;
			}
			else
			{
				ListObjArray[i].Obj.className='TempletItem';
				ListObjArray[i].Selected=false;
			}
		}
	}
}
function AddSiteFolder()
{
	location='?Action=Addsitefolder';
}
function AddSite()
{
	location='?Action=Addsite';
}

function Lock(Falg)
{
	var SelectedSite='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.SiteID!=null)
			{
				if (SelectedSite=='') SelectedSite=ListObjArray[i].Obj.SiteID;
				else  SelectedSite=SelectedSite+'***'+ListObjArray[i].Obj.SiteID;
			}
		}
	}
	if (SelectedSite!='')
	{
		if (Falg) location='?Action=Lock&LockID='+SelectedSite;
		else location='?Action=UNLock&LockID='+SelectedSite;
	}
	else alert('请选择站点');
}
function EditSite()
{
	var SelectedSite='',SelectedSiteFolder='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.SiteID!=null)
			{
				if (SelectedSite=='') SelectedSite=ListObjArray[i].Obj.SiteID;
				else  SelectedSite=SelectedSite+'***'+ListObjArray[i].Obj.SiteID;
			}
			if (ListObjArray[i].Obj.SiteFolderID!=null)
			{
				if (SelectedSiteFolder=='') SelectedSiteFolder=ListObjArray[i].Obj.SiteFolderID;
				else  SelectedSiteFolder=SelectedSiteFolder+'***'+ListObjArray[i].Obj.SiteFolderID;
			}
		}
	}
	if (SelectedSite!='')
	{
		SelectedSiteFolder='';
		if (SelectedSite.indexOf('***')==-1) location='SitemodifyOne.asp?SiteID='+SelectedSite;
		else alert('请选择一个站点');
	}
	else if (SelectedSiteFolder!='')
	{
		if (SelectedSiteFolder.indexOf('***')==-1)
		location='site.asp?Action=Addsitefolder&SiteFolderID='+SelectedSiteFolder;
		else alert('请选择一个目录');
	}
	else
	{
		alert('请选择站点或者栏目');
	} 

}
function EditSiteGuide()
{
	var SelectedSite='',SelectedSiteFolder='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.SiteID!=null)
			{
				if (SelectedSite=='') SelectedSite=ListObjArray[i].Obj.SiteID;
				else  SelectedSite=SelectedSite+'***'+ListObjArray[i].Obj.SiteID;
			}
			else if (ListObjArray[i].Obj.SiteFolderID!=null)
			{
				if (SelectedSiteFolder=='') SelectedSiteFolder=ListObjArray[i].Obj.SiteFolderID;
				else  SelectedSiteFolder=SelectedSiteFolder+'***'+ListObjArray[i].Obj.SiteFolderID;
			}
		}
	}
	if (SelectedSite!='')
	{
		if (SelectedSite.indexOf('***')==-1) location='Sitemodify.asp?SiteID='+SelectedSite;
		else alert('请选择一个站点');
	}
	else alert('请选择一个站点');
}
function DelSite()
{
	var SelectedSite='',SelectedSiteFolder='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.SiteID!=null)
			{
				if (SelectedSite=='') SelectedSite=ListObjArray[i].Obj.SiteID;
				else  SelectedSite=SelectedSite+'***'+ListObjArray[i].Obj.SiteID;
			}
			if (ListObjArray[i].Obj.SiteFolderID!=null)
			{
				if (SelectedSiteFolder=='') SelectedSiteFolder=ListObjArray[i].Obj.SiteFolderID;
				else  SelectedSiteFolder=SelectedSiteFolder+'***'+ListObjArray[i].Obj.SiteFolderID;
			}
		}
	}
	if (SelectedSite!='' || SelectedSiteFolder!='')
	{
		if (confirm('确定要删除吗？')==true)
		window.location='?action=Del&Id='+SelectedSite+'&SiteFolderID='+SelectedSiteFolder;
	}
	else alert('请选择站点或栏目');
}
function CopySite()
{
	var SelectedSite='',SelectedSiteFolder='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.SiteID!=null)
			{
				if (SelectedSite=='') SelectedSite=ListObjArray[i].Obj.SiteID;
				else  SelectedSite=SelectedSite+'***'+ListObjArray[i].Obj.SiteID;
			}
			if (ListObjArray[i].Obj.SiteFolderID!=null)
			{
				if (SelectedSiteFolder=='') SelectedSiteFolder=ListObjArray[i].Obj.SiteFolderID;
				else  SelectedSiteFolder=SelectedSiteFolder+'***'+ListObjArray[i].Obj.SiteFolderID;
			}
		}
	}
	if (SelectedSite!='' || SelectedSiteFolder!='')
	{
		location='Site.asp?vs=Copy&SiteID='+SelectedSite+'&SiteFolderID='+SelectedSiteFolder;
	}
	else alert('请选择站点或者栏目');
}
function StartCollect()
{
	var SelectedSite='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.SiteID!=null)
			{
				if (ListObjArray[i].Obj.IsCollect=='True')
				{
					if (SelectedSite=='') SelectedSite=ListObjArray[i].Obj.SiteID;
					else  SelectedSite=SelectedSite+'***'+ListObjArray[i].Obj.SiteID;
				}
			}
		}
	}
	if (SelectedSite!='')
	{
		Num = InsertScript();
	//	alert(Num);
		if (Num!='back'&& Num!='0')
		{
			if (Num==""||Num==null)
			{
				Num="allNews"
			}
			location='Collecting.asp?SiteID='+SelectedSite+'&Num='+Num;
		}
	}
	else alert('请选择站点，或者您设置有误！');
}
function StartOneSiteCollect(ID)
{
	Num = InsertScript();
	//alert(Num);
	if (Num!='back'&& Num!='0')
	{
		if (Num==""||Num==null)
		{
			Num="allNews"
		}
		location='Collecting.asp?SiteID='+ID+'&Num='+Num;
	}
}
function StartCollectAtServer()
{
	var SelectedSite='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.SiteID!=null)
			{
				if (ListObjArray[i].Obj.IsCollect=='True')
				{
					if (SelectedSite=='') SelectedSite=ListObjArray[i].Obj.SiteID;
					else  SelectedSite=SelectedSite+'***'+ListObjArray[i].Obj.SiteID;
				}
			}
		}
	}
	if (SelectedSite!='')
	{
		if (confirm(CopyRightStr))
		{
			var ObjStr="Microsoft.XMLHTTP"
			var HTTP=new ActiveXObject(ObjStr);
			HTTP.onreadystatechange=function()
			{
				if (HTTP.readyState == 4) // 调用完毕
				{
				alert(HTTP.status);
					if (HTTP.status == 200) // 加载成功
					{
						alert (HTTP.responseText);
					}
				}
			} 
			HTTP.open("get",'CollectingAtServer.asp?SiteID='+SelectedSite,true);
			HTTP.send();
			HTTP=null;
		}
	}
	else alert('请选择站点');
}

function ResumeCollect()
{
	var SelectedSite='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.SiteID!=null)
			{
				if (ListObjArray[i].Obj.IsCollect=='True')
				{
					if (SelectedSite=='') SelectedSite=ListObjArray[i].Obj.SiteID;
					else  SelectedSite=SelectedSite+'***'+ListObjArray[i].Obj.SiteID;
				}
			}
		}
	}
	if (SelectedSite!='')
	{
		if(confirm("      欢迎使用风讯新闻采集系统，如果设计到版权问题与风讯科技发展有限公司无关，确定要继续采集吗？"))location='Collecting.asp?CollectType=ResumeCollect&SiteID='+SelectedSite;		
	}
	else alert('请选择站点');
}

function Export()
{
	var SelectedSite='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.SiteID!=null)
			{
				if (SelectedSite=='') SelectedSite=ListObjArray[i].Obj.SiteID;
				else  SelectedSite=SelectedSite+'***'+ListObjArray[i].Obj.SiteID;
			}
		}
	}
	if (SelectedSite!='')
	{
		location='Export.asp?ID='+SelectedSite;
	}
	else alert('请选择站点');
}
function Import()
{
	location='Import.asp';
}

function InsertScript()
{
	var ReturnValue='';
	ReturnValue=showModalDialog("NewsNum.asp",window,'dialogWidth:260pt;dialogHeight:120pt;status:no;help:no;scroll:no;');
	return ReturnValue;
}
function ChangeFolder(Obj)
{
	location.href='site.asp?Action=SubFolder&FolderID='+Obj.SiteFolderID;
}
function backTop()
{
	location.href = 'site.asp'
}
</script>