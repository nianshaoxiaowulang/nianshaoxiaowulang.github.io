<% Option Explicit %>
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../Refresh/RefreshFunction.asp" -->
<!--#include file="../Refresh/SelectFunction.asp" -->
<%
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
Dim DBC,Conn,RecordConn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
DBC.ConnStr = "DBQ=" + server.mappath(RecordDataBaseConnectStr) + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
Set RecordConn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../Refresh/Function.asp" -->
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P070607") then Call ReturnError1()
Dim SaveFilePath,FSOObj,PromptInfo,TempletFileName,FromDate,TentDate,RsRecordObj
Dim DateDiffNum,Refresh_LoopVar,RefreshTime
FromDate = Request("FromDate")
TentDate = Request("TentDate")
PromptInfo = ""
if SysRootDir = "" then
	SaveFilePath = "/" & RecordNewsListSavePath & "/"
	TempletFileName = "/" & TempletDir
else
	SaveFilePath = "/" & SysRootDir & "/" & RecordNewsListSavePath & "/"
	TempletFileName = "/" & SysRootDir & "/" & TempletDir
end if
if FromDate <> "" then
	SetRefreshValue "Record",""
	GetAvailableDoMain
	TempletFileName = Server.MapPath(TempletFileName) & "\File.htm"
	Set FSOObj = Server.CreateObject(G_FS_FSO)
	if FSOObj.FileExists(TempletFileName) = False then
		PromptInfo = "�鵵ģ��File.htm�����ڣ�����ӹ鵵ģ��������ɣ�"
	else
		if TentDate <> "" then
			DateDiffNum = DateDiff("d",FromDate,TentDate)
		else
			DateDiffNum = 0
		end if
		for Refresh_LoopVar = 0 to DateDiffNum
			RefreshTime = DateAdd("d",Refresh_LoopVar,FromDate)
			Set RsRecordObj = RecordConn.Execute("Select * from FS_News where DateDiff('d',FileTime,#" & RefreshTime & "#)=0 order by ID Desc")
			if Refresh_LoopVar Mod 4 = 0 then PromptInfo = PromptInfo & "<br>"
			PromptInfo = PromptInfo & RefreshRecord(SaveFilePath,TempletFileName,FSOObj)
			RsRecordObj.Close
			Set RsRecordObj = Nothing
		Next
	end if
	Set FSOObj = Nothing
else
	PromptInfo = "û��ѡ��ʱ��"
end if
Call PromptFunction
Function RefreshRecord(SaveFilePath,TempletFileName,FSOObj)
	Dim FileStreamObj,FileContent,FileObj,SaveFileName
	'On Error Resume Next
	Set FileObj = FSOObj.GetFile(TempletFileName)
	Set FileStreamObj = FileObj.OpenAsTextStream(1)
	SaveFileName = SaveFilePath & RefreshTime & ".htm"
	if Not FileStreamObj.AtEndOfStream then
		FileContent = FileStreamObj.ReadAll
		FileContent = ReplaceAllServerFlag(ReplaceAllLable(FileContent))
		RefreshRecord = "<a target=""_blank"" href=""" & SaveFileName & """>" & RefreshTime & "</a>&nbsp;&nbsp;&nbsp;&nbsp;"
	else
		RefreshRecord = "ģ������Ϊ��"
	end if
	Select Case AvailableRefreshType
		Case 0
			FSOSaveFile FileContent,SaveFileName
		Case 1
			SaveFile FileContent,SaveFileName
		Case Else
			FSOSaveFile FileContent,SaveFileName
	End Select
	Set FileStreamObj = Nothing
End Function
Sub PromptFunction()
	Set Conn = Nothing
	Set RecordConn = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�鵵�����б����ɹ���</title>
</head>
<link rel="stylesheet" href="../../../CSS/FS_css.css">
<body topmargin="0" leftmargin="0" oncontextmenu="return false;">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="28" class="ButtonListLeft">
<div align="center"><strong>�鵵�����б����ɹ���</strong></div></td>
</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td><div align="center"><font color="#FF0000">������Ϣ:</font></div></td>
  </tr>
  <tr> 
    <td><div align="center"> 
        <% = PromptInfo %>
      </div></td>
  </tr>
</table>
</body>
</html>
<%
End Sub
%>
