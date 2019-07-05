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
if Not JudgePopedomTF(Session("Name"),"P030805") then Call ReturnError()
Dim RsBackObj,RsLableObj,LableID,SQLStr
LableID = Request("LableID")
SQLStr = "Select * From FS_LableBackUp Where ID=" &LableID
Set RsBackObj = Server.CreateObject(G_FS_RS)
RsBackObj.Open SQLStr,Conn,3,3
Set RsLableObj = Server.CreateObject(G_FS_RS)
RsLableObj.Open "Select * From FS_Lable Where ID="&LableID,conn,3,3
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
	<% if RsLableObj.eof  then %>
	<tr>
	<td width="25%" height="60" rowspan="2">
		<div align="center"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
	<td align="center" height="15"><font size=2.5>此标签已被删除，请指定还原的位置：</font>	  </td></tr>
	<tr>
	  <td align="center" height="15"><select name="TypeID">
        <option value="0">根类型</option>
		<%
		Dim TempObj
		Set TempObj = conn.Execute("select * from FS_LableType")
		do while not TempObj.eof 
		%>
		<option <% if CStr(RsBackObj("Type")) = Cstr(TempObj("ID")) then response.Write("Selected")%> value="<% = TempObj("ID") %>"><% = TempObj("TypeName") %></option>
		<%
		TempObj.MoveNext
		loop
		%>
      </select></td>
    </tr>
	<tr>
   	   <td height="30" align="center" colspan="2"> 
	   	<input type=hidden name=operation value=Modify>
	    <input type="submit" value="  确定  " class=Anbutc> 
		<input type="reset" value="  取消  " onclick="window.close()" class=Anbutc> </td>
	</tr>
	<% else %>
	<tr>
	<td width="25%" height="60">
		<div align="center"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
	<td align="center" height="30"><font size=2.5>此标签存在，确定还原？</font></td></tr>
	<tr>
   	   <td height="30" align="center" colspan="2"> 
	   	<input type=hidden name=operationChild value=ModifyChild>
	    <input type="submit" value="  确定  " class=Anbutc> 
		<input type="reset" value="  取消  " onclick="window.close()" class=Anbutc> </td>
	</tr>
	<% end if %>
</form>
</table>
</body>
</html>
<%
	if Request.Form("operation") = "Modify" then
		RsLableObj.Close
		RsLableObj.Open "Select * From FS_Lable",conn,3,3
		RsLableObj.AddNew
		RsLableObj("Type") = Request.Form("TypeID")
		Revert()
	end if
	if Request.Form("operationChild") = "ModifyChild" then
		RsLableObj("Type") = RsBackObj("Type")
		Revert()
	end if
Function Revert()		
	On Error Resume Next
	RsLableObj("LableName") = RsBackObj("LableName")
	RsLableObj("LableContent") = RsBackObj("LableContent")
	RsLableObj("Description") = RsBackObj("Description")
	RsLableObj.Update
	RsLableObj.Close
	RsLableObj.Open "Select ID From FS_Lable Where LableName='"&  RsBackObj("LableName") &"'",conn,1,1
	RsBackObj("ID") = RsLableObj("ID")
	RsBackObj.Update
	RsLableObj.Close
	if err>0 then
		%>
		<script language="javascript">
			alert('还原失败');
			window.close();
		</script>
		<%
	else
		%>
		<script language="javascript">
			window.close();
		</script>
		<%
	end if
End Function
RsBackObj.Close
Set RsBackObj = Nothing
Set conn = Nothing
%>
