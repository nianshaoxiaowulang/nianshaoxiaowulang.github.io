<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../Inc/Cls_Info.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
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
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<!--#include file="../../../Inc/Session.asp" -->
<%
if Not (JudgePopedomTF(Session("Name"),"P020320")) then Call ReturnError()
Dim SourceSpecial,TargetSpecial,MergeSql,RsTarGetObj,SourceDir,FSO,DelSource
SourceSpecial = Request("SourceSpecial")
TargetSpecial = Request.form("TargetSpecial")
DelSource=Request.form("DelSource")
If TargetSpecial<>"" and TargetSpecial<>SourceSpecial then 
	Set RsTarGetObj = Conn.Execute("Select SaveFilePath,EName from FS_Special where SpecialID='" & SourceSpecial & "'")

	if SysRootDir = "" then
		SourceDir = RsTarGetObj("SaveFilePath") & "/" & RsTarGetObj("EName")
	else
		SourceDir = "/" & SysRootDir & RsTarGetObj("SaveFilePath") & "/" & RsTarGetObj("EName")
	end if
	SourceDir = Server.MapPath(SourceDir)
	Set FSO = Server.CreateObject(G_FS_FSO)
	If FSO.FolderExists(SourceDir) then
		FSO.DeleteFolder SourceDir
	End if
	Set RsTarGetObj=Nothing
	Set FSO = Nothing
	dim RsSpecialID,TempSpeID
	Set RsSpecialID=Server.CreateObject(G_FS_RS)
	MergeSql = "select Newsid,SpecialID from FS_News where SpecialID like '%" & SourceSpecial & "%'"
	RsSpecialID.open MergeSql,Conn,3,3
	
			Do while not RsSpecialID.eof
				If instr(1,RsSpecialID(1),TargetSpecial)=0 then
					TempSpeID=","&RsSpecialID(1)&","
					TempSpeID=replace(TempSpeID, SourceSpecial & ",",TargetSpecial&",")
					TempSpeID=mid(TempSpeID,2,len(TempSpeID)-2)
					conn.execute("update FS_news set SpecialID='"& TempSpeID &"' where Newsid='"&RsSpecialID(0)&"'")
				End If
				RsSpecialID.update
				RsSpecialID.movenext
			loop
	If DelSource="DelSource" then 
		MergeSql="Delete from FS_Special where SpecialID='" & SourceSpecial & "'"
		Conn.Execute(MergeSql)
	end if
	Response.Write("<script>window.close();</script>")
elseif TargetSpecial=SourceSpecial then 
	Response.Write("<script>alert('源专题和目标专题一样，不可以合并！');window.close();</script>")
	Response.end
else
	Dim TempClassListStr
	TempClassListStr=SpecialClassIDList(SourceSpecial)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
	<head>
	<link href="../../../CSS/ModeWindow.css" rel="stylesheet" type="text/css">
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<title>专题合并</title>
	</head>
	<body leftmargin="0" topmargin="0">
	  <form action="?SourceSpecial=<%=SourceSpecial%>" method="post" name="ClassForm">
	  <table width="100%">
	  <tr height="30" valign="bottom">
		<td width="70%" align="right">选择你要合并到的专题
		</td>
		<td align="left"><select name="TargetSpecial">
		<% =TempClassListStr %>
		</select>
		</td>
		</tr>
		<tr height="30">
		<td width="70%" align="right">同时删除被合并的专题</td>
		<td align="left"><input name="DelSource"  type="CheckBox" value="DelSource"></td>
	  </tr>
		<tr>
		<td align="center" colspan="2">
		<input name="NumClass"  type="submit" id="NumClass" value="确 定">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name="CloseOk"  type="button" id="NumClass" value="关 闭" onClick="window.close();">
		  </td>
	  </tr>
	  </table>
	  </form>
	</body>
	</html>
<%
End if
Function SpecialClassIDList(SpecialID)
	Dim TempRs,TempStr
	Set TempRs = Conn.Execute("Select SpecialID,CName from FS_Special")
	do while Not TempRs.Eof
		If SpecialID<>TempRs("SpecialID") then '不显示被合并的专题
			SpecialClassIDList = SpecialClassIDList & "<option value="&TempRs("SpecialID") & ">" & TempRs("CName") & chr(13)
		End if
		TempRs.MoveNext
	loop
	TempRs.Close
	Set TempRs = Nothing
End Function
%>
