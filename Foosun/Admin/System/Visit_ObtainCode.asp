<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
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
if Not JudgePopedomTF(Session("Name"),"P080501") then Call ReturnError1()
Dim TruePlusDir
If PlusDir="" then
	TruePlusDir=""
Else
	TruePlusDir="/"&PlusDir
End If

%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����ͳ�ƴ���</title>
</head>
<body topmargin="2" leftmargin="2" oncontextmenu="//return false;">
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td height="28" class="ButtonListLeft"> <div align="center"><strong>����ͳ�ƴ������</strong></div></td>
  </tr>
</table>
<br>
<br>
<table width="85%" height="90"  border="0" align="center" cellpadding="3" cellspacing="1" bordercolor="e6e6e6" bgcolor="dddddd">
  <tr bgcolor="#FFFFFF"> 
    <td width="16%" valign="middle"> 
      <div align="center">��ͼ��</div></td>
    <td width="84%" valign="middle"><SPAN class=small2><FONT face="Verdana, Arial, Helvetica, sans-serif">&lt;script 
      src="<%=confimsn("DoMain")%><%=TruePlusDir%>/count/count.asp?Type=Pic"&gt;&lt;/script&gt;</FONT></SPAN></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td valign="middle"> 
      <div align="center">��ͼ��</div></td>
    <td valign="middle"><FONT face="Verdana, Arial, Helvetica, sans-serif"><SPAN class=small2>&lt;script 
      src="<%=confimsn("DoMain")%><%=TruePlusDir%>/count/count.asp"&gt;&lt;/script&gt;</SPAN></FONT></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td valign="middle"> 
      <div align="center">����ͳ��</div></td>
    <td valign="middle"><FONT face="Verdana, Arial, Helvetica, sans-serif"><SPAN class=small2>&lt;script 
      src="<%=confimsn("DoMain")%><%=TruePlusDir%>/count/count.asp?Type=Word"&gt;&lt;/script&gt;</SPAN></FONT></td>
  </tr>
</table>
<div align="center"></div>
</body>
</html>
