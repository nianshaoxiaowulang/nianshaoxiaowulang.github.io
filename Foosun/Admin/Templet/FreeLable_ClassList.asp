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
'产品目录：风讯产品N系列
'软件名称：风讯网站信息管理系统
'当前版本：Foosun Content Manager System V1.0.0
'最新更新：2004.8
'==============================================================================
'Copyright (C) 2002-2004 Foosun.Net  All rights reserved.
'商业注册联系：028-85098980-601,技术支持：028-85098980-606、607,客户支持：608
'产品咨询QQ：159410,655071 
'技术支持QQ：66252421 
'程序开发：风讯开发组 & 风讯插件开发组
'Email:service@cooin.com
'论坛支持：风讯在线论坛(http://bbs.cooin.com   http://bbs.foosun.net)
'官方网站：www.Foosun.net  演示站点：www.cooin.com    开发者园地：www.aspsun.cn
'==============================================================================
'免费版本请在新闻首页保留版权信息，并做上本站LOGO友情连接
'风讯在线保留此程序的法律追究权利
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
<title>无标题文档</title>
<style>
Td{Font size:12Px;}
</style>
</head>
<body leftmargin="10" topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
      
    <td width="25%" height="20"><strong>中文名称：</strong></td>
      
    <td width="25%"><strong>编号：</strong></td>
    <td width="25%"><strong>中文名称：</strong></td>
    <td width="25%"><strong>编号：</strong></td>
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
