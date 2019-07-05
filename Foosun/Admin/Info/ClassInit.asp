<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../Inc/Cls_Info.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<%
Dim DBC,Conn,sRootDir
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
if SysRootDir<>"" then sRootDir="/"+SysRootDir else sRootDir=""
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
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P010513") then Call ReturnError()
if Not JudgePopedomTF(Session("Name"),""&Request("ClassID")&"") then Call ReturnError1()
Dim ClassID,MyFile,AlertInfo,RsClassEditObj,Sql,DelNewsSysRootDir
Set MyFile=Server.CreateObject(G_FS_FSO)
ClassID = Request("ClassID")
If Request("TrueDel")="TrueDel" then 
	If SysRootDir<>"" then 
		DelNewsSysRootDir="/" & SysRootDir
	Else
		DelNewsSysRootDir=""
	End If
	if ClassID <> "" then
		Sql = "Select * from FS_NewsClass where ClassID='" & ClassID & "' and DelFlag=0"
		Set RsClassEditObj = Conn.Execute(Sql)
		if RsClassEditObj.Eof then
	'		Set RsClassEditObj = Nothing
	'		Set Conn = Nothing
			AlertInfo="栏目已经被删除 "
		else		
			Sql = "Delete from FS_News where ClassID='" & ClassID & "'"
			Conn.Execute(Sql)

			if Err.Number <> 0 then AlertInfo= "删除栏目下的新闻失败":err.clear
			Sql = "Delete from FS_Contribution where ClassID='" & ClassID & "'"
			Conn.Execute(Sql)
			if Err.Number <> 0 then AlertInfo= "删除栏目下的投稿失败":err.clear
			Sql = "Delete from FS_DownLoad where ClassID='" & ClassID & "'"
			Conn.Execute(Sql)
			if Err.Number <> 0 then AlertInfo= "删除栏目下的下载失败":err.clear
	
			If MyFile.FolderExists(Server.Mappath(DelNewsSysRootDir&RsClassEditObj("SaveFilePath")&"/"&RsClassEditObj("ClassEName"))) then
				MyFile.DeleteFolder(Server.Mappath(DelNewsSysRootDir&RsClassEditObj("SaveFilePath")&"/"&RsClassEditObj("ClassEName")))
			End if
			If Err.Number <> 0 then AlertInfo=AlertInfo & "删除栏目中的新闻文件失败 "
		end if
		If AlertInfo="" then AlertInfo="初始化完成！ "
	else
		AlertInfo="参数传递错误！ "
	end if
	%>
	<script>
		alert('初始化完成！');
		window.close();
	</script>
	<%
Else
	ShowTrueInfo
End If
Sub ShowTrueInfo
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
	<head>
	<link href="../../../CSS/ModeWindow.css" rel="stylesheet" type="text/css">
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<title>栏目初始化</title>
	</head>
	<body leftmargin="0" topmargin="0">
	  <form action="?ClassID=<%=ClassID%>" method="post" name="ClassForm">
	  <table width="100%">
	  <tr height="20">
		<td width="70%" align="right">
		</td>
		</tr>
		<tr height="30" align="center">
		<td width="70%" align="center">初始化后，栏目中的所有新闻、投稿、下载都将被删除！确认？</td>
	  </tr>
		<tr>
		<td align="center">
		<input name="TrueDel"  type="hidden" value="TrueDel">
		<input name="NumClass"  type="submit" value="确 定">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name="CloseOk"  type="button" value="取 消" onClick="window.close();">
		  </td>
	  </tr>
	  </table>
	  </form>
	</body>
	</html>
<%
End Sub
%>
