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
if Not JudgePopedomTF(Session("Name"),"P030902") then Call ReturnError()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>修改标签</title>
</head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet">
<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0">
<br>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<form name=form method=post action="" >
	<tr>
	<td width="25%" height="60">
		<div align="right"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
	<td align="center" height="30"><font size=2.5>确定删除吗？</font></td></tr>
	<tr>
   	   <td height="30" align="center" colspan="2"><input type="hidden" name="ID" value="<% = Request("ID") %>"> 
        <input type=hidden name=operation value=Modify>
	    <input type="submit" value="  确定  " class=Anbutc> 
		<input type="reset" value="  取消  " onclick="window.close()" class=Anbutc> </td>
	</tr>
</form>
</table>
</body>
</html>
<%
if Request.Form("operation") = "Modify" then
	On Error Resume Next
	Dim LableID,SQLStr
	LableID = Replace(request("ID"),"***",",")
	SQLStr = "Delete From FS_LableBackUp where ID in (" & LableID & ")"
	Conn.Execute(SQLStr)
	if Err.Number = 0 then
		%>
		<script language="javascript">
		dialogArguments.location.reload();
		window.close();
		</script>
		<%
	else
		%>
		<script language="javascript">
		alert ("有错误发生，请重试")
		dialogArguments.location.reload();
		window.close();
		</script>
		<%
	end if
end if
Set Conn = Nothing
%>
