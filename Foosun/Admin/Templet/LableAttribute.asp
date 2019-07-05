<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P030901") then Call ReturnError()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>标签属性</title>
<link rel="stylesheet" href="../../../CSS/ModeWindow.css">
</head>
<%
Dim ID
ID = Request("ID")
%>
<body topmargin="0" leftmargin="0" scroll=no>
<table width="100%" height="100" border="0" cellpadding="0" cellspacing="0">
  <tr>
	<td><iframe id="Editer" src="LableContent.asp?ID=<% = ID %>" scrolling="yes" width="100%" height="200" frameborder="1"></iframe></td>
</tr>
<tr>
	<td height="30">
<div align="center">
        
      </div></td>
</tr>
</table>
</body>
</html>
<%
Set Conn = Nothing
%>