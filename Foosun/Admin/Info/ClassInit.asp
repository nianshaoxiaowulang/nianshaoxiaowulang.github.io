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
'������ƣ���Ѷ��վ��Ϣ����ϵͳ
'��ǰ�汾��Foosun Content Manager System(FoosunCMS V3.1.0930)
'���¸��£�2005.10
'==============================================================================
'Copyright (C) 2002-2004 Foosun.Net  All rights reserved.
'��ҵע����ϵ��028-85098980-601,��Ŀ������028-85098980-606��609,�ͻ�֧�֣�608
'��Ʒ��ѯQQ��394226379,159410,125114015
'����֧��QQ��315485710,66252421 
'��Ŀ����QQ��415637671��655071
'���򿪷����Ĵ���Ѷ�Ƽ���չ���޹�˾(Foosun Inc.)
'Email:service@Foosun.cn
'MSN��skoolls@hotmail.com
'��̳֧�֣���Ѷ������̳(http://bbs.foosun.net)
'�ٷ���վ��www.Foosun.cn  ��ʾվ�㣺test.cooin.com 
'��վͨϵ��(���ܿ��ٽ�վϵ��)��www.ewebs.cn
'==============================================================================
'��Ѱ汾���ڳ�����ҳ������Ȩ��Ϣ�������ϱ�վLOGO��������
'��Ѷ��˾�����˳���ķ���׷��Ȩ��
'�������2�ο��������뾭����Ѷ��˾������������׷����������
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
			AlertInfo="��Ŀ�Ѿ���ɾ�� "
		else		
			Sql = "Delete from FS_News where ClassID='" & ClassID & "'"
			Conn.Execute(Sql)

			if Err.Number <> 0 then AlertInfo= "ɾ����Ŀ�µ�����ʧ��":err.clear
			Sql = "Delete from FS_Contribution where ClassID='" & ClassID & "'"
			Conn.Execute(Sql)
			if Err.Number <> 0 then AlertInfo= "ɾ����Ŀ�µ�Ͷ��ʧ��":err.clear
			Sql = "Delete from FS_DownLoad where ClassID='" & ClassID & "'"
			Conn.Execute(Sql)
			if Err.Number <> 0 then AlertInfo= "ɾ����Ŀ�µ�����ʧ��":err.clear
	
			If MyFile.FolderExists(Server.Mappath(DelNewsSysRootDir&RsClassEditObj("SaveFilePath")&"/"&RsClassEditObj("ClassEName"))) then
				MyFile.DeleteFolder(Server.Mappath(DelNewsSysRootDir&RsClassEditObj("SaveFilePath")&"/"&RsClassEditObj("ClassEName")))
			End if
			If Err.Number <> 0 then AlertInfo=AlertInfo & "ɾ����Ŀ�е������ļ�ʧ�� "
		end if
		If AlertInfo="" then AlertInfo="��ʼ����ɣ� "
	else
		AlertInfo="�������ݴ��� "
	end if
	%>
	<script>
		alert('��ʼ����ɣ�');
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
	<title>��Ŀ��ʼ��</title>
	</head>
	<body leftmargin="0" topmargin="0">
	  <form action="?ClassID=<%=ClassID%>" method="post" name="ClassForm">
	  <table width="100%">
	  <tr height="20">
		<td width="70%" align="right">
		</td>
		</tr>
		<tr height="30" align="center">
		<td width="70%" align="center">��ʼ������Ŀ�е��������š�Ͷ�塢���ض�����ɾ����ȷ�ϣ�</td>
	  </tr>
		<tr>
		<td align="center">
		<input name="TrueDel"  type="hidden" value="TrueDel">
		<input name="NumClass"  type="submit" value="ȷ ��">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name="CloseOk"  type="button" value="ȡ ��" onClick="window.close();">
		  </td>
	  </tr>
	  </table>
	  </form>
	</body>
	</html>
<%
End Sub
%>
