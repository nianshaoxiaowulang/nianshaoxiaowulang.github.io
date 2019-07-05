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
if Not JudgePopedomTF(Session("Name"),"P031303") then Call ReturnError()
Dim DelType,DelLable
DelLable = Request("DelLable")
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>修改自由标签</title>
</head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet">
<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0">
<br>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<form name=form method=post action="" >
	<tr>
	<td width="25%" height="60">
		<div align="right"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
	<td height="30"><font size=2.5>是否删除这项？</font></td></tr>
	<tr>
   	   <td height="30" align="center" colspan="2">
        <input type="hidden" name="DelLable" value="<% = DelLable %>"> 
        <input name=Action type=hidden id="Action" value=Submit>
	    <input type="submit" value="  确定  " class=Anbutc> 
		<input type="reset" value="  取消  " onclick="window.close()" class=Anbutc> </td>
	</tr>
</form>
</table>
</body>
</html>
<%
if Request.Form("Action")="Submit" then
	if DelLable <> "" then
		DelLable = Replace(DelLable,"***","','")
		Conn.Execute("Delete from FS_FreeLable where FreeLableID in ('" & DelLable & "')")
	end if
	if err.Number=0 then
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
