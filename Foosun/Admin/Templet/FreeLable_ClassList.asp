<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
Dim  DBC,Conn,TempClassListStr,TempListStr
Set  DBC = New DataBaseClass
Set  Conn = DBC.OpenConnection()
Set  DBC = Nothing
'==============================================================================
'��ƷĿ¼����Ѷ��ƷNϵ��
'������ƣ���Ѷ��վ��Ϣ����ϵͳ
'��ǰ�汾��Foosun Content Manager System V1.0.0
'���¸��£�2004.8
'==============================================================================
'Copyright (C) 2002-2004 Foosun.Net  All rights reserved.
'��ҵע����ϵ��028-85098980-601,����֧�֣�028-85098980-606��607,�ͻ�֧�֣�608
'��Ʒ��ѯQQ��159410,655071 
'����֧��QQ��66252421 
'���򿪷�����Ѷ������ & ��Ѷ���������
'Email:service@cooin.com
'��̳֧�֣���Ѷ������̳(http://bbs.cooin.com   http://bbs.foosun.net)
'�ٷ���վ��www.Foosun.net  ��ʾվ�㣺www.cooin.com    ������԰�أ�www.aspsun.cn
'==============================================================================
'��Ѱ汾����������ҳ������Ȩ��Ϣ�������ϱ�վLOGO��������
'��Ѷ���߱����˳���ķ���׷��Ȩ��
'==============================================================================
%>
<!--#include file="../../../Inc/Session.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P030802") and Not JudgePopedomTF(Session("Name"),"P030803")  then
 	Call ReturnError1()
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ޱ����ĵ�</title>
<style>
Td{Font size:12Px;}
</style>
</head>
<body leftmargin="10" topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
      
    <td width="25%" height="20"><strong>�������ƣ�</strong></td>
      
    <td width="25%"><strong>��ţ�</strong></td>
    <td width="25%"><strong>�������ƣ�</strong></td>
    <td width="25%"><strong>��ţ�</strong></td>
    </tr>
    <%
	Dim Rs
	Set Rs=Conn.execute("Select ClassID,ClassCName From FS_NewsClass Order by ClassCName")
	Do while Not Rs.Eof
	%>
    <tr> 
      <td height="15"><font color="#0066FF"><%=Rs("ClassCName")%></font></td>
      <td><font color="#0066FF"><%=Rs("ClassID")%></font></td>
	<%
		Rs.Movenext
		If Not Rs.Eof Then
	%>
      <td><font color="#0066FF"><%=Rs("ClassCName")%></font></td>
      <td><font color="#0066FF"><%=Rs("ClassID")%></font></td>
	<%
		Rs.Movenext
		else
	%>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
 	<%
		end if
	%>
  </tr>
  <tr>
  	<td colspan=4 height=2><hr></td>
  </tr>
    <%
	Loop
	Set Rs = Nothing
	%>
  </table>
</body>
</html>
