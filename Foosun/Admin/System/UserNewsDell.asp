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
'if Not ((JudgePopedomTF(Session("Name"),"P040403")) OR (JudgePopedomTF(Session("Name"),"P040405"))) then Call ReturnError()
dim UserIDP,UserDellObj,OperateType,TipStr
if Request("ID")<>"" and Request("OperateType")<>"" then
	UserIDP = Request("ID")
	OperateType = Cstr(Request("OperateType"))
	If OperateType = "Dell" then
		'if Not JudgePopedomTF(Session("Name"),"P040403") then Call ReturnError()
		TipStr = "ɾ��"
	ElseIf OperateType = "Lock" then
		'if Not JudgePopedomTF(Session("Name"),"P040405") then Call ReturnError()
		TipStr = "�������"
	Else 
		'if Not JudgePopedomTF(Session("Name"),"P040405") then Call ReturnError()
		TipStr = "����"
	End If
else
	Response.Write("<script>alert(""�������ݴ���"");window.close();</script>")
	Response.End
end if 
%>
<html>
<head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��Աɾ��</title>
</head>
<body leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="5" cellpadding="0">
<form action="" name="JSDellForm" method="post">
  <tr> 
    <td width="7%" height="10">&nbsp;</td>
    <td width="28%" rowspan="3"><div align="center"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
    <td width="59%">&nbsp;</td>
    <td width="6%" height="10">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>��ȷ��Ҫ<%=TipStr%>��?</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td height="2">&nbsp;</td>
    <td height="2">&nbsp;</td>
    <td height="2">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td colspan="2"><div align="center"> 
        <input type="submit" name="Submit" value=" ȷ �� ">
        <input type="hidden" name="action" value="Submit">
        <input type="button" name="Submit2" value=" ȡ �� " onClick="window.close();">
      </div></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td height="10">&nbsp;</td>
    <td height="10" colspan="2">&nbsp;</td>
    <td height="10">&nbsp;</td>
  </tr>
</form>
</table>
</body>
</html>
<%
if Request.Form("action")="Submit" then
	UserIDP = Replace(UserIDP,"***",",")
	If OperateType = "Dell" then
		'if Not JudgePopedomTF(Session("Name"),"P040403") then Call ReturnError()
		Conn.Execute("Delete from FS_MemberNews where ID in (" & UserIDP & ")")
	Elseif OperateType = "isLock" then
		'if Not JudgePopedomTF(Session("Name"),"P040405") then Call ReturnError()
		Conn.Execute("Update FS_MemberNews set isLock=1 where ID in (" & UserIDP & ")")
	Else
		'if Not JudgePopedomTF(Session("Name"),"P040405") then Call ReturnError()
		Conn.Execute("Update FS_MemberNews set isLock=0 where ID in (" & UserIDP & ")")
	End If
	Response.Write("<script>dialogArguments.location.reload();window.close();</script>")
	Response.End
end if
%>