<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
<html>
<head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>
<%
Dim DBC,Conn,RsJsObj,ID,SQLStr
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
Set ID = Request("ID")
%>
<!--#include file="../../../Inc/Session.asp" -->

<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P060705") then Call ReturnError()
%>
<body topmargin="0" leftmargin="0" scroll=no>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <%
Dim Str
if ID = "" then
	Str = "参数传递错误"
else
	SQLStr = "Select Info from FS_FreeJS Where ID="&ID
	Set RsJsObj=Conn.Execute(SQLStr)
	if Not RsJsObj.Eof then
		Str = RsJsObj("Info")
	else
		Str = "参数传递错误"
	end if
end if
%>
  <tr>
    <td nowrap><textarea readonly name="textarea" rows="11" style="width:100%;"><% = Str %></textarea></td>
  </tr>
  <tr> 
    <td height="30" nowrap>
<div align="center">
        <input name="Submitdd" type="button" id="Submitdd" onClick="window.close();" value=" 关 闭 ">
      </div></td>
  </tr>
</table>
</body>
</html>
<%
Set RsJsObj = Nothing
Set Conn = Nothing
%>
