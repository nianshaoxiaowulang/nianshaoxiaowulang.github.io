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
if Not JudgePopedomTF(Session("Name"),"P020310") then Call ReturnError()
Dim SpecialID,MyFile,AlertInfo,RsClassEditObj,Sql,DelNewsSysRootDir,RsSpecialID
Set MyFile=Server.CreateObject(G_FS_FSO)
SpecialID = Request("SpecialID")
If Request("TrueDel")="TrueDel" then 
	If SysRootDir<>"" then 
		DelNewsSysRootDir="/" & SysRootDir
	Else
		DelNewsSysRootDir=""
	End If
	
	if SpecialID <> "" then
		Set RsClassEditObj= Server.CreateObject(G_FS_RS)
		RsClassEditObj.Open "Select * from FS_Special where SpecialID='" & SpecialID&"'",Conn,3,3
		if RsClassEditObj.Eof then
			AlertInfo="ר���Ѿ���ɾ�� "
		else
			Sql="Select SpecialID from FS_News"
			Set RsSpecialID=Server.CreateObject(G_FS_RS)
			RsSpecialID.Open Sql,Conn,3,3
			Dim TempSpeID
			Do while not RsSpecialID.eof
				If instr(1,RsSpecialID(0),SpecialID)>0 then
					If instr(1,RsSpecialID(0),",")>0 then
						TempSpeID=","&RsSpecialID(0)&","
						TempSpeID=replace(TempSpeID, SpecialID & ",","")
						TempSpeID=mid(TempSpeID,2,len(TempSpeID)-2)
						RsSpecialID("SpecialID")=TempSpeID
					Else
						RsSpecialID("SpecialID")=""
					End If
				End If
				RsSpecialID.update
				RsSpecialID.movenext
			loop
			If MyFile.FileExists(Server.Mappath(DelNewsSysRootDir&RsClassEditObj("SaveFilePath")&"/"&RsClassEditObj("EName")&"/index."&RsClassEditObj("FileExtName"))) then
				MyFile.DeleteFile(Server.Mappath(DelNewsSysRootDir&RsClassEditObj("SaveFilePath")&"/"&RsClassEditObj("EName")&"/index."&RsClassEditObj("FileExtName")))
			End if
			Set RsClassEditObj=Nothing
			Set MyFile=Nothing
		end if
	else
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
	<title>ר���ʼ��</title>
	</head>
	<body leftmargin="0" topmargin="0">
	  <form action="?SpecialID=<%=SpecialID%>" method="post" name="ClassForm">
	  <table width="100%">
	  <tr height="20" valign="bottom">
		<td width="70%" align="right">
		</td>
		</tr>
		<tr height="30">
		<td width="70%" align="center">��ʼ����ר�⽫�ָ��ս���ʱ��״̬��ȷ�ϣ�</td>
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

