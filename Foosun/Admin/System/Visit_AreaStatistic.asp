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
if Not JudgePopedomTF(Session("Name"),"P080508") then Call ReturnError1()
%>
<html>
<head>
<title>�����ߵ���ͳ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script src="../JS/PublicJS.js" language="JavaScript"></script>
<body bgcolor="#FFFFFF" topmargin="2" leftmargin="2">
<table height="26" width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
      <td height="28" class="ButtonListLeft"><div align="center"><strong>�����ߵ���ͳ��</strong></div></td>
  </tr>
</table>
<%
Dim RsAreaObj,Sql
Set RsAreaObj = Server.CreateObject(G_FS_RS)
Sql="Select Area From FS_FlowStatistic"
RsAreaObj.Open Sql,Conn,3,3
Dim AreaType
Dim NumIn,NumOut,NumOther
NumIn=0
NumOut=0
NumOther=0
Do While not RsAreaObj.Eof
	AreaType= RsAreaObj("Area")
	Select Case AreaType
	Case "�������ڲ���"
		NumIn=NumIn+1
	Case "δ֪����"
		NumOther=NumOther+1
	Case Else
		NumOut=NumOut+1
	End Select
	RsAreaObj.MoveNext
Loop
%>
<%
Dim AllNum
AllNum=NumIn+NumOut+NumOther
%>
<table width=96% border=0 cellpadding=2>
	<tr>
		<td align=center>�����ߵ���ͳ��ͼ��</td>
	</tr>
	<tr>
		<td align=center>
			<table align=center>
        <tr valign=bottom >
					
          <td nowap>�ڲ���</td>
					
          <td align=left background=../../Images/Visit/tu_back_2.gif valign=middle width=150><img src=../../Images/Visit/bar2.gif height=15></td>
					<td nowap><% =NumIn %></td>
				</tr>
				<tr valign=bottom >
					
          <td nowap>�ⲿ��</td>
					
          <td align=left background=../../Images/Visit/tu_back_2.gif valign=middle width=150><img src=../../Images/Visit/bar2.gif height=15></td>
					<td nowap><% =NumOut %></td>
				</tr>
				<tr valign=bottom >
					
          <td align="right" nowap>δ֪</td>
					
          <td align=left background=../../Images/Visit/tu_back_2.gif valign=middle width=150><img src=../../Images/Visit/bar2.gif height=15></td>
					<td nowap><% =NumOther %></td>
				</tr>
				<tr valign=cente>
					<td align=right nowap>��</td>
					
          <td align=left background=../../Images/Visit/tu_back_2.gif valign=middle width=150><img src=../../Images/Visit/bar2.gif width="150" height=15></td>
					<td nowap><% = AllNum %></td>
			</table><br>
		</td>
	</tr>
</table>