<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
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
if Not JudgePopedomTF(Session("Name"),"P040404") then Call ReturnError()
	Dim UGSetObj,UserIDP
	UserIDP = Request("ID")
	Set UGSetObj = Conn.Execute("Select GroupID from FS_Members where ID="&Clng(UserIDP)&"")
	If UGSetObj.eof then
		Response.Write("<script>alert(""�������ݴ���"");window.close();</script>")
		Response.End
	End If
%>
<html>
<head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>���û�Ա��</title>
</head>
<body leftmargin="0" topmargin="0">
<form action="" method="post" name="UGDellForm">
<table width="100%"  border="0" cellspacing="5" cellpadding="0">
  <tr>
    <td width="9%">&nbsp;</td>
    <td width="14%">&nbsp;</td>
    <td width="77%">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>��Ա��</td>
    <td>
	<select name="GroupID" style="width:90%">
	<option value="0" <%If UGSetObj("GroupID")="0" or UGSetObj("GroupID")="" then Response.Write("selected") end if%>> </option>
    <%
	Dim ChooseGroupObj
	Set ChooseGroupObj = Conn.Execute("Select Name,GroupID from FS_MemGroup order by AddTime desc")
	do while not ChooseGroupObj.eof
	%>
	<option value="<%=ChooseGroupObj("GroupID")%>" <%If Cstr(UGSetObj("GroupID")) = Cstr(ChooseGroupObj("GroupID")) then Response.Write("selected") end if%>><%=ChooseGroupObj("Name")%></option>
	<%
	ChooseGroupObj.MoveNext
	Loop
	ChooseGroupObj.Close
	Set ChooseGroupObj = Nothing
	%>
	</select></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td colspan="3"><div align="center">
      <input type="submit" name="Submit" value=" ȷ �� ">
      <input name="action" type="hidden" id="action" value="trues">
      <input type="button" name="Submit2" value=" ȡ �� " onClick="window.close();">
    </div></td>
  </tr>
</table>
</form>
</body>
</html>
<%
If Request.Form("action") = "trues" then
	Conn.Execute("Update FS_Members set GroupID='"&Cstr(Request.Form("GroupID"))&"' where ID="&UserIDP&"")
	Set Conn = Nothing
	Response.Write("<script>dialogArguments.location.reload();window.close();</script>")
	Response.End
End If
Conn.Close
Set Conn = Nothing
%>