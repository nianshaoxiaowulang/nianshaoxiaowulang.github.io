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
	Response.Write("<script>alert('Դר���Ŀ��ר��һ���������Ժϲ���');window.close();</script>")
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
	<title>ר��ϲ�</title>
	</head>
	<body leftmargin="0" topmargin="0">
	  <form action="?SourceSpecial=<%=SourceSpecial%>" method="post" name="ClassForm">
	  <table width="100%">
	  <tr height="30" valign="bottom">
		<td width="70%" align="right">ѡ����Ҫ�ϲ�����ר��
		</td>
		<td align="left"><select name="TargetSpecial">
		<% =TempClassListStr %>
		</select>
		</td>
		</tr>
		<tr height="30">
		<td width="70%" align="right">ͬʱɾ�����ϲ���ר��</td>
		<td align="left"><input name="DelSource"  type="CheckBox" value="DelSource"></td>
	  </tr>
		<tr>
		<td align="center" colspan="2">
		<input name="NumClass"  type="submit" id="NumClass" value="ȷ ��">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name="CloseOk"  type="button" id="NumClass" value="�� ��" onClick="window.close();">
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
		If SpecialID<>TempRs("SpecialID") then '����ʾ���ϲ���ר��
			SpecialClassIDList = SpecialClassIDList & "<option value="&TempRs("SpecialID") & ">" & TempRs("CName") & chr(13)
		End if
		TempRs.MoveNext
	loop
	TempRs.Close
	Set TempRs = Nothing
End Function
%>
