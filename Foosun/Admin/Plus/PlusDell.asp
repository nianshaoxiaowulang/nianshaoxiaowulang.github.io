<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
Dim PlusID,Types,TypeStr
PlusID = Request("PlusID")
Types = Request("Types")
If Types = "Shows" then
	if Not JudgePopedomTF(Session("Name"),"P080604") then Call ReturnError()
	TypeStr = "显示"
elseif Types = "Hide" then
	if Not JudgePopedomTF(Session("Name"),"P080605") then Call ReturnError()
	TypeStr = "隐藏"
else
	if Not JudgePopedomTF(Session("Name"),"P080603") then Call ReturnError()
	TypeStr = "删除"
end if
%>
<html>
<head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>插件管理</title>
</head>
<body topmargin="0" leftmargin="0" ondragstart="return false;" onselectstart="return false;">
<form name="PlusDelForm" method="post" action="">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="5">
  <tr>
    <td width="32%" rowspan="3"><div align="center"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
    <td width="68%">&nbsp;</td>
  </tr>
  <tr>
    <td>您确定要<%=TypeStr%>吗?</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2"><div align="center">
      <input type="submit" name="Submit" value=" 确 定 ">
      <input name="action" type="hidden" id="action" value="trues">
      <input type="button" name="Submit2" value=" 取 消 " onclick="window.close();">
    </div></td>
  </tr>
  <tr>
    <td colspan="2">&nbsp;</td>
  </tr>
</table>
</form>
</body>
</html>
<%
If Request.Form("action") = "trues" then
	if PlusID <> "" then
		PlusID = Replace(PlusID,"***",",")
		If Types = "Shows" then
			if Not JudgePopedomTF(Session("Name"),"P080604") then Call ReturnError()
			Conn.Execute("Update FS_Plus set ShowTF=1 where ID in (" & PlusID & ")")
		elseif Types = "Hide" then
			if Not JudgePopedomTF(Session("Name"),"P080605") then Call ReturnError()
			Conn.Execute("Update FS_Plus set ShowTF=0 where ID in (" & PlusID & ")")
		else
			if Not JudgePopedomTF(Session("Name"),"P080603") then Call ReturnError()
			Conn.Execute("Delete from FS_Plus where ID in (" & PlusID & ")")
		end if
	end if
	Response.Write("<script>dialogArguments.location.reload();window.close();</script>")
End If
Set Conn = Nothing
%>