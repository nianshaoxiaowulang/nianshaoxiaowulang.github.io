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
if Not ((JudgePopedomTF(Session("Name"),"P070303")) OR (JudgePopedomTF(Session("Name"),"P070304")) OR (JudgePopedomTF(Session("Name"),"P070305"))) then Call ReturnError()
Dim VoteID,Types,TempStr
VoteID = Cstr(Request("VoteID"))
Types = Cstr(Request("Types"))
if Types = "Dell" then
	if Not JudgePopedomTF(Session("Name"),"P070303") then Call ReturnError()
	TempStr = "删除"
elseif Types = "Open" then
	if Not JudgePopedomTF(Session("Name"),"P070304") then Call ReturnError()
	TempStr = "开启"
elseif Types = "Close" then
	if Not JudgePopedomTF(Session("Name"),"P070305") then Call ReturnError()
	TempStr = "关闭"
end if
'%>
<html>
<head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>删除投票项目</title>
</head>
<body leftmargin="0" topmargin="0">
<form name="VoteDelForm" action="" method="post">
<table width="100%" border="0" cellspacing="5" cellpadding="0">
  <tr> 
    <td width="7%" height="10">&nbsp;</td>
    <td width="17%" rowspan="3"><div align="center"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
    <td width="70%">&nbsp;</td>
    <td width="6%" height="10">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>您确定要<%=TempStr%>该投票项目?</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td height="2">&nbsp;</td>
    <td height="2">&nbsp;</td>
    <td height="2">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td colspan="2"><div align="center"> 
        <input type="submit" name="Submit" value=" 确 定 ">
        <input type="hidden" name="action" value="trues">
        <input type="button" name="Submit2" value=" 取 消 " onClick="window.close();">
      </div></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td height="10">&nbsp;</td>
    <td height="10" colspan="2">&nbsp;</td>
    <td height="10">&nbsp;</td>
  </tr>
</table>
</form>
</body>
</html>

<%
If Request.Form("action") = "trues" then
	if VoteID <> "" then
		VoteID = Replace(VoteID,"***","','")
		if Types = "Dell" then
			if Not JudgePopedomTF(Session("Name"),"P070303") then Call ReturnError()
			Conn.Execute("Delete from FS_VoteOption where VoteID in ('" & VoteID & "')")
			Conn.Execute("Delete from FS_Vote where VoteID in ('" & VoteID & "')")
		end if
		if Types = "Open" then
			if Not JudgePopedomTF(Session("Name"),"P070304") then Call ReturnError()
			Conn.Execute("Update FS_Vote set State=1 where VoteID in ('" & VoteID & "')")
		end if
		if Types = "Close" then
			if Not JudgePopedomTF(Session("Name"),"P070305") then Call ReturnError()
			Conn.Execute("Update FS_Vote set State=0 where VoteID in ('" & VoteID & "')")
		end if
	end if
	Response.Write("<script>dialogArguments.location.reload();window.close();</script>")
	Response.End
End If
%>