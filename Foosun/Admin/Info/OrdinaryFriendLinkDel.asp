<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
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
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P070503") then Call ReturnError()
Dim FriendLinkID
FriendLinkID = Request("FriendLinkID")
%>
<html>
<head>
<link rel="stylesheet" href="../../../CSS/ModeWindow.css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�������ӹ���</title>
</head>
<body topmargin="0" leftmargin="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="3">
  <form method="post" name="DelForm">
    <tr> 
      <td><div align="center"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
      <td colspan="2">��ȷ��Ҫɾ������������?</td>
    </tr>
    <tr> 
      <td colspan="3"><div align="center"> 
          <input type="submit" name="Submit" value=" ȷ �� ">
          <input type="hidden" name="Action" value="Submit">
          <input type="hidden" name="FriendLinkID" value="<% = FriendLinkID %>">
          <input type="button" onClick="window.close();" name="Submit2" value=" ȡ �� " >
      </div></td>
    </tr>
  </form>
</table>
</body>
</html>
<%
if Request.Form("Action") = "Submit" then
	Conn.Execute("Delete from FS_FriendLink where ID in (" & Replace(FriendLinkID,"***",",") & ")")
	Response.Write("<script>dialogArguments.location.reload();window.close();</script>")
end if
Set Conn = Nothing
%>
