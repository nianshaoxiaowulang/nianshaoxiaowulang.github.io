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
if Not JudgePopedomTF(Session("Name"),"P040303") then Call ReturnError()
Dim UserGroupID
UserGroupID = Request("ID")
%>
<html>
<head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��Ա��ɾ��</title>
</head>
<body leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="5" cellpadding="0">
<form action="" name="JSDellForm" method="post">
  <tr> 
    <td width="7%" height="10"></td>
    <td width="19%" rowspan="3"><div align="center"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
    <td width="68%"></td>
    <td width="6%" height="10"></td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>��ȷ��Ҫɾ���˻�Ա��<br>��ɾ���˻�Ա���µ����л�Ա?</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td height="2"></td>
    <td height="2"></td>
    <td height="2"></td>
  </tr>
  <tr> 
    <td></td>
    <td colspan="2"><div align="center"> 
        <input type="submit" name="Submit" value=" ȷ �� ">
        <input type="hidden" name="action" value="Submit">
        <input type="hidden" name="ID" value="<% = UserGroupID %>">
        <input type="button" name="Submit2" value=" ȡ �� " onClick="window.close();">
      </div></td>
    <td></td>
  </tr>
  <tr> 
    <td height="10"></td>
    <td height="10" colspan="2"></td>
    <td height="10"></td>
  </tr>
</form>
</table>
</body>
</html>
<%
if Request.Form("action") = "Submit" then
	UserGroupID = Replace(UserGroupID,"***",",")
	Conn.Execute("Delete from FS_Message where MeRead in (Select MemName from FS_Members where GroupID in (Select GroupID from FS_MemGroup where ID in (" & UserGroupID & ")))")
	Conn.Execute("Delete from FS_Members where GroupID in (Select GroupID from FS_MemGroup where ID in (" & UserGroupID & "))")
	Conn.Execute("Delete from FS_MemGroup where ID in (" & UserGroupID & ")")
	Response.write("<script>dialogArguments.location.reload();window.close();</script>")
	Response.end
end if
%>