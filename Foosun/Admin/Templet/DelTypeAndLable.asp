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
if Not JudgePopedomTF(Session("Name"),"P030804") then Call ReturnError()
Dim DelType,DelLable
DelType = Request("DelType")
DelLable = Request("DelLable")
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
	<td align="center" height="30"><font size=2.5>是否删除这项？</font></td></tr>
	<tr>
   	   <td height="30" align="center" colspan="2"><input type="hidden" name="DelType" value="<% = DelType %>">
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
	Dim TempArray,LoopVar,DelAllTypeID,TempDelAllTypeID
	TempArray = Split(DelType,"***")
	for LoopVar = LBound(TempArray) to UBound(TempArray)
		if TempArray(LoopVar) <> "" then
			TempDelAllTypeID = TempArray(LoopVar) & ChildTypeIDList(TempArray(LoopVar))
		end if
		if TempDelAllTypeID <> "" then
			if DelAllTypeID = "" then
				DelAllTypeID = TempDelAllTypeID
			else
				DelAllTypeID = DelAllTypeID & "," & TempDelAllTypeID
			end if
		end if
		TempDelAllTypeID = ""
	Next
	if DelAllTypeID <> "" then
		Conn.Execute("Delete from FS_LableType where ID in (" & DelAllTypeID & ")")
		Conn.Execute("Delete from FS_Lable where Type in (" & DelAllTypeID & ")")
	end if
	if DelLable <> "" then
		DelLable = Replace(DelLable,"***",",")
		Conn.Execute("Delete from FS_Lable where ID in (" & DelLable & ")")
	end if
	if err.Number=0 then
		%>
			<script language="javascript">
			window.close();
			</script>
		<%
	else
		%>
			<script language="javascript">
			alert ("有错误发生，请重试")
			window.close();
			</script>
		<%
	end if
end if
Set Conn = Nothing

Function ChildTypeIDList(TypeID)
	Dim TempRs,TempStr
	Set TempRs = Conn.Execute("Select ID from FS_LableType where ParentID=" & TypeID)
	do while Not TempRs.Eof
		ChildTypeIDList = ChildTypeIDList & "," & TempRs("ID") & ""
		ChildTypeIDList = ChildTypeIDList & ChildTypeIDList(TempRs("ID"))
		TempRs.MoveNext
	loop
	TempRs.Close
	Set TempRs = Nothing
End Function
%>
