<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../Inc/Cls_Info.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../refresh/Function.asp" -->
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
if Not (JudgePopedomTF(Session("Name"),"P010512")) then Call ReturnError()
if Not JudgePopedomTF(Session("Name"),"" & Request("SourceClass") & "") then Call ReturnError()
if Not JudgePopedomTF(Session("Name"),"" & Request("ObjectClass") & "") then Call ReturnError()
Dim ShowSubmitTF,SourceClass,ObjectClass,Result,ShowStr,RsSoueceObj,RsObjectObj,AllowOperation,AllClassID
SourceClass = Request("SourceClass")
ObjectClass = Request("ObjectClass")
AllClassID = "'" & SourceClass & "'" & ChildClassIDList(SourceClass)
If SourceClass<>ObjectClass then 
	Result = Request("Result")
	ShowSubmitTF = True
	AllowOperation = True
	if (SourceClass = "") OR (ObjectClass = "") then
		ShowSubmitTF = False
		ShowStr = "�������ݴ���"
	else
		Set RsSoueceObj = Conn.Execute("Select * from FS_NewsClass where ClassID='" & SourceClass & "'")
		if RsSoueceObj.Eof then
			ShowSubmitTF = False
			ShowStr = "Դ��Ŀ������"
			AllowOperation = False
		else
			ShowStr = "ȷ��Ҫ�ѣ�" & RsSoueceObj("ClassCName") & "�ݺϲ���"
		end if
		Set RsObjectObj = Conn.Execute("Select * from FS_NewsClass where ClassID='" & ObjectClass & "'")
		if RsObjectObj.Eof then
			ShowSubmitTF = False
			ShowStr = "Ŀ����Ŀ������"
			AllowOperation = False
		else
			if AllowOperation = True then
				ShowStr = ShowStr & "��" & RsObjectObj("ClassCName") & "����"
			end if
		end if
		if Result = "Submit" then
			if AllowOperation = True then
				Dim MergeSql
				MoveNewsFile "",SourceClass,ObjectClass
				MergeSql = "Update FS_News Set ClassID='" & ObjectClass & "' where ClassID in(" & AllClassID & ")"
				Conn.Execute(MergeSql)
				MergeSql = "Update FS_download Set ClassID='" & ObjectClass & "' where ClassID in(" & SourceClass & ")"
				Conn.Execute(MergeSql)
				MergeSql = "Update FS_Contribution set ClassID='" & ObjectClass & "' where ClassID in (" & AllClassID & ")"
				'------------/l
				DelClass(SourceClass)
				'---------------
				Set RsSoueceObj = Nothing
				Set RsObjectObj = Nothing
				Response.write("<script>window.close();</script>")
				Response.end
			else
				Set RsSoueceObj = Nothing
				Set RsObjectObj = Nothing
				Response.write("<script>alert('�ϲ���Ŀ������');window.close();</script>")
				Response.end
			end if
		end if
		Set RsSoueceObj = Nothing
		Set RsObjectObj = Nothing
	end if
Else
	ShowStr="Դ��Ŀ��Ŀ����Ŀ��ͬ�������Ժϲ���"
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ƶ����߿���������Ŀ</title>
</head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet">
<body topmargin="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <form name="OperateForm" action="" method="post">
    <tr> 
      <td height="20" colspan="2"></td>
    </tr>
    <tr>
      <td width="28%" rowspan="3"><div align="center"><strong><font size="2"><img src="../../Images/Question.gif" width="39" height="37"></font></strong></div></td>
    </tr>
    <tr>
      <td width="72%"><% = ShowStr %></td>
    </tr>
    <tr>
      <td width="72%" height="20"><font color="#FF0000">�ϲ���ɾ��ԭ����Ŀ(������������Ŀ)</font>
<input name="DelSource" type="checkbox" value="Del"></td>
    </tr>
    <tr> 
<%
if ShowSubmitTF = true then
%>
      <td colspan="2" height="40"><div align="center"> 
          <input type="submit" name="Submit" value=" ȷ �� ">
          <input name="Result" type="hidden" id="Result" value="Submit">
          <input type="button" name="Submit2" onClick="window.close();" value=" ȡ �� ">
        </div></td>
<%
end if
%>
    </tr>
  </form>
</table>
</body>
</html>
<%
Set Conn = Nothing
Sub DelClass(DelClassID)

	Dim AllClassID,Sql,DelNewsSysRootDir,MyFile
	AllClassID = "'" & DelClassID & "'" & ChildClassIDList(DelClassID)
	If SysRootDir<>"" then 
		DelNewsSysRootDir="/"& SysRootDir
	else
		DelNewsSysRootDir=""
	End If
	Set MyFile=Server.CreateObject(G_FS_FSO)
	'---------------------�����ļ�ɾ��-------------------------------------
	Dim DelClassFileObj
	Set DelClassFileObj = Conn.Execute("Select ClassEName,SaveFilePath from FS_NewsClass where ClassID in ("&AllClassID&")")
	Do while Not DelClassFileObj.eof
		If MyFile.FolderExists(Server.Mappath(DelNewsSysRootDir&DelClassFileObj("SaveFilePath")&"/"&DelClassFileObj("ClassEName"))) then
			MyFile.DeleteFolder(Server.Mappath(DelNewsSysRootDir&DelClassFileObj("SaveFilePath")&"/"&DelClassFileObj("ClassEName")))
		End if
		DelClassFileObj.MoveNext
	Loop
	DelClassFileObj.Close
	Set DelClassFileObj = Nothing
	set MyFile=Nothing
	'����������������������������������������
	'Sql = "Delete from News where ClassID in (" & AllClassID & ")"
	'Conn.Execute(Sql)
	'if Err.Number <> 0 then Alert "ɾ����Ŀ�µ�����ʧ��"
	'Sql = "Delete from Contribution where ClassID in (" & AllClassID & ")"
	'Conn.Execute(Sql)
	'if Err.Number <> 0 then Alert "ɾ����Ŀ�µ�Ͷ��ʧ��"
	'Sql = "Delete from DownLoad where ClassID in (" & AllClassID & ")"
	'Conn.Execute(Sql)
	'if Err.Number <> 0 then Alert "ɾ����Ŀ�µ�����ʧ��"
	If request("DelSource")="Del" then 
		Sql = "Delete from FS_NewsClass where ClassID in (" & AllClassID & ")"
		Conn.Execute(Sql)
		if Err.Number <> 0 then Alert "ɾ����Ŀʧ��"
	End if
End Sub
Function ChildClassIDList(ClassID)
	Dim TempRs,TempStr
	Set TempRs = Conn.Execute("Select ClassID from FS_NewsClass where ParentID = '" & ClassID & "'")
	do while Not TempRs.Eof
		ChildClassIDList = ChildClassIDList & ",'" & TempRs("ClassID") & "'"
		ChildClassIDList = ChildClassIDList & ChildClassIDList(TempRs("ClassID"))
		TempRs.MoveNext
	loop
	TempRs.Close
	Set TempRs = Nothing
End Function
%>
