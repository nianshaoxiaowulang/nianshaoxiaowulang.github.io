<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
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
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not ((JudgePopedomTF(Session("Name"),"P040203")) OR (JudgePopedomTF(Session("Name"),"P040206"))) then Call ReturnError()
Dim AdminID,Result,OperateType,PromptText
PromptText = ""
OperateType = Request("OperateType")
Result = Request("Result")
AdminID = Replace(Replace(Request("AdminID"),"'",""),"""","")
if OperateType = "DelAdmin" then
	if Not JudgePopedomTF(Session("Name"),"P040203") then Call ReturnError()
	PromptText = "ȷ��Ҫɾ���˹���Ա��"
elseif OperateType = "LockAdmin" then
	if Not JudgePopedomTF(Session("Name"),"P040206") then Call ReturnError()
	PromptText = "ȷ��Ҫ�����˹���Ա��"
elseif OperateType = "UNLockAdmin" then
	if Not JudgePopedomTF(Session("Name"),"P040206") then Call ReturnError()
	PromptText = "ȷ��Ҫ�����˹���Ա��"
else
	PromptText = ""
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ɾ��������������Ա</title>
</head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet">
<body  oncontextmenu="return false;">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <form name="OperateForm" action="" method="post">
  <tr> 
      <td width="7%" height="20">&nbsp;</td>
      <td width="27%" rowspan="3"><div align="center"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
      <td width="66%">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><% = PromptText %></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td height="20">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="3"><div align="center"><input type="submit" name="Submit" value=" ȷ �� ">
          <input name="OperateType" value="<% = OperateType %>" type="hidden" id="OperateType">
          <input name="AdminID" value="<% = AdminID %>" type="hidden" id="AdminID">
          <input name="Result" type="hidden" id="Result" value="Submit">
          <input type="button" name="Submit2" onClick="dialogArguments.location.reload();window.close();" value=" ȡ �� ">
      </div></td>
    </tr>
 </form>
</table>
</body>
</html>
<%
if Result = "Submit" then
	Dim ReturnCheckInfo
	AdminID = Replace(AdminID,"***",",")
	if OperateType = "DelAdmin" then
		if Not JudgePopedomTF(Session("Name"),"P040203") then Call ReturnError()
		Conn.Execute("delete from FS_Admin where ID in (" & AdminID & ") and GroupID<>0")
	elseif OperateType = "LockAdmin" then
		if Not JudgePopedomTF(Session("Name"),"P040206") then Call ReturnError()
		Conn.Execute("update FS_Admin set Lock=1 where ID in (" & AdminID & ") and GroupID<>0")
	elseif OperateType = "UNLockAdmin" then
		if Not JudgePopedomTF(Session("Name"),"P040206") then Call ReturnError()
		Conn.Execute("update FS_Admin set Lock=0 where ID in (" & AdminID & ") and GroupID<>0")
	end if
	if Err.Number = 0 then
		%>
		<script language="JavaScript">
		dialogArguments.location.reload();
		window.close();
		</script>
		<%
	else
		%>
		<script language="JavaScript">
		alert('��������');
		dialogArguments.location.reload();
		window.close();
		</script>
		<%
	end if
end if
%>