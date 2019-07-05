<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
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
%><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
</head>
<%
Dim ID,RsLableObj,SQLStr
ID = Request("ID")
%>
<body topmargin="8" leftmargin="8">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <%
if ID = "" then
%>
  <tr> 
    <td>参数传递错误</td>
  </tr>
  <%
else
	Set RsLableObj = Server.CreateObject(G_FS_RS)
	SQLStr="Select * From FS_Lable Where ID=" & ID
	Set RsLableObj = Conn.Execute(SQLStr)
	if Not RsLableObj.Eof then
%>
  <tr> 
    <td>描述： <% = RsLableObj("Description") %></td>
  </tr>
  <tr> 
    <td>内容： <% = RsLableObj("LableContent") %></td>
  </tr>
  <%
	else
%>
  <tr> 
    <td>标签不存在</td>
  </tr>
  <%
	end if
%>
  <tr> 
    <td height="30">&nbsp;</td>
  </tr>
  <%
end if
%>
</table>
</body>
</html>
<%
Set RsLableObj = Nothing
Set Conn = Nothing
%>
<script language="JavaScript">
document.designMode="On";
document.open();
</script>