<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="inc/Config.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
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
Dim DBC,Conn,CollectConn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
DBC.ConnStr = CollectDBConnectionStr
Set CollectConn = DBC.OpenConnection()
Set DBC = Nothing
'�ж�Ȩ��
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P080302") then Call ReturnError()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ޱ����ĵ�</title>
</head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet">
<body topmargin="0" leftmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <form action="" method="pose" name="Form">
    <tr> 
      <td width="120" height="80"> 
        <div align="center"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
      <td height="80"> 
        <div align="left">ȷ��Ҫɾ���� 
          <input type="hidden" value="Submit" name="Action">
          <input type="hidden" name="NewsIDStr" value="<% = Request("NewsIDStr") %>">
        </div></td>
    </tr>
    <tr> 
      <td colspan="2"><div align="center"> 
          <input type="submit" name="Submit" value=" ȷ �� ">
          <input type="button" name="Submit2" onClick="window.close();" value=" ȡ �� ">
        </div></td>
    </tr>
  </form>
</table>
</body>
</html>
<%
if Request("Action") = "Submit" then
	Dim NewsIDStr,DelSql
	NewsIDStr = Request("NewsIDStr")
	if NewsIDStr <> "" then
		'On Error Resume Next
		NewsIDStr = Replace(NewsIDStr,"***",",")
		DelSql = "Delete from FS_News where ID in (" & NewsIDStr & ")"
		CollectConn.Execute(DelSql)
		if Err.Number = 0 then
			Response.Write("<script>dialogArguments.location.reload();window.close();</script>")
		else
			Response.Write("<script>dialogArguments.location.reload();window.close();</script>")
		end if
		Set CollectConn = Nothing
	end if
end if
%>