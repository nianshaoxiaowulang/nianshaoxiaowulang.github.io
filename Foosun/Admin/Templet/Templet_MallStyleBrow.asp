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
if Not JudgePopedomTF(Session("Name"),"P031203") then Call ReturnError()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>查看下载列表样式</title>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
</head>
<%
Dim ID,RsLableObj,SQLStr
ID = Request("ID")
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<body topmargin="0" leftmargin="0">
<table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
  <%
if ID = "" then
%>
  <tr> 
    <td bgcolor="#FFFFFF">参数传递错误</td>
  </tr>
  <%
else
	Set RsLableObj = Server.CreateObject(G_FS_RS)
	SQLStr="Select * From FS_MallListStyle Where ID=" & ID 
	Set RsLableObj = Conn.Execute(SQLStr)
	if Not RsLableObj.Eof then
%>
  <tr> 
    <td bgcolor="#FFFFFF">名称：
<% = RsLableObj("Name") %></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF">内容：
<% = RsLableObj("Content") %></td>
  </tr>
  <%
	Else
%>
  <tr> 
    <td bgcolor="#FFFFFF">标签不存在</td>
  </tr>
  <%
	End if
End if
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