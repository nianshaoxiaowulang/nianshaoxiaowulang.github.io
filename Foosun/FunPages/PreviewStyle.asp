<% Option Explicit %>
<!--#include file="../../Inc/Const.asp" -->
<!--#include file="../../Inc/Cls_DB.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../Inc/CheckPopedom.asp" -->
<!--#include file="../../Inc/Session.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P030800") then Call ReturnError()

Dim FreeLableID,StyleContent,SqlStr,Rs
FreeLableID = Replace(Request("FreeLableID"),"'","")
If FreeLableID = "" Then
	StyleContent = "自由标签预览"
Else
	SqlStr = "select StyleContent from FS_freelable where freelableid = '"&FreeLableID&"'"
	Set Rs = conn.Execute(SqlStr)
	If Rs.eof Then
		StyleContent = "无效的自由标签编号"
	Else
		StyleContent = Rs("StyleContent")
	End if
End if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>自由标签预览</title>
<link rel="stylesheet" href="../../CSS/Style.css">
</head>
<body topmargin="0" leftmargin="0">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
  	<td align="center" valign="middle"><%=StyleContent%>
	</td>
  </tr>
</table>
</body>
</html>
