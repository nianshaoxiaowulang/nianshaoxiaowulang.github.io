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
if Not JudgePopedomTF(Session("Name"),"P030903") then Call ReturnError()
Dim RsBackObj,RsLableObj,LableID,SQLStr,AlreadyBackUpLable,PromptStr
LableID = Replace(Request("BackUpLable"),"***",",")
Set RsBackObj = Server.CreateObject(G_FS_RS)
RsBackObj.Open "Select * From FS_LableBackUp Where ID in (" & LableID & ")",Conn,3,3
do while Not RsBackObj.Eof
	if AlreadyBackUpLable = "" then
		AlreadyBackUpLable = RsBackObj("LableName")
	else
		AlreadyBackUpLable = AlreadyBackUpLable & "|" & RsBackObj("LableName")
	end if
	RsBackObj.MoveNext
Loop
if AlreadyBackUpLable = "" then
	PromptStr = "确定要备份吗？"
else
	PromptStr = AlreadyBackUpLable & "已经备份。<br>确定备份吗？"
end if
Set RsBackObj = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>修改标签</title>
</head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet">
<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0" scroll=no>
<br>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <form name=LableForm method=post action="" >
	<tr>
	<td width="25%" height="60">
		<div align="center"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
	<td align="center" height="30"><font size=2.5><% = PromptStr %></font></td></tr>
	<tr>
   	   <td height="30" align="center" colspan="2">
	   	<input type="hidden" name="Operation">
		<% if AlreadyBackUpLable <> "" then %>
	    <input name="按钮" type="button" value=" 覆 盖 " onclick="SubmitForm('Cover');">
        <input name="按钮" type="button" value="不 覆 盖" onclick="SubmitForm('Submit');">
		<% else %>
        <input name="按钮" type="button" value=" 确 定 " onclick="SubmitForm('Submit');">
		<% end if %>
        <input type="button" value="  取消  " onclick="window.close()" class=Anbutc> </td>
	</tr>
  </form>
</table>
</body>
</html>
<script>
function SubmitForm(Value)
{
	document.LableForm.Operation.value=Value;
	document.LableForm.submit();
}
</script>
<%
Dim Operation
Operation = Request.Form("operation")
if Operation <> "" then
	'On Error Resume Next
	SQLStr = "Select * From FS_Lable Where ID in (" & LableID & ")"
	Set RsLableObj = Server.CreateObject(G_FS_RS)
	RsLableObj.Open SQLStr,Conn,1,1
	do while Not RsLableObj.Eof
		if Operation = "Submit" then
			Back RsLableObj("ID"),false
		elseif Operation = "Cover" then
			Back RsLableObj("ID"),true
		end if
		RsLableObj.MoveNext
	Loop
	Set RsLableObj = Nothing
	Set Conn = Nothing
	Response.write("<script>window.close();</script>")
end if
Sub Back(BackUpLableID,CoverTF)
	Dim RsCheckBackObj
	Set RsCheckBackObj = Server.CreateObject(G_FS_RS)
	RsCheckBackObj.Open "Select * from FS_LableBackUp where ID=" & BackUpLableID,Conn,3,3
	if Not RsCheckBackObj.Eof then
		if CoverTF = true then
			RsCheckBackObj("LableName") = RsLableObj("LableName")
			RsCheckBackObj("LableContent") = RsLableObj("LableContent")
			RsCheckBackObj("Description") = RsLableObj("Description")
			RsCheckBackObj("Type") = RsLableObj("Type")
			RsCheckBackObj("BackUpTime") = now()
			RsCheckBackObj.Update
		end if
	else
		RsCheckBackObj.AddNew()
		RsCheckBackObj("ID") = RsLableObj("ID")
		RsCheckBackObj("LableName") = RsLableObj("LableName")
		RsCheckBackObj("LableContent") = RsLableObj("LableContent")
		RsCheckBackObj("Description") = RsLableObj("Description")
		RsCheckBackObj("Type") = RsLableObj("Type")
		RsCheckBackObj("BackUpTime") = now()
		RsCheckBackObj.Update
	end if
	Set RsCheckBackObj = Nothing
End Sub
%>