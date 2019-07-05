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
if Not JudgePopedomTF(Session("Name"),"P0501") then Call ReturnError()
if Not ((JudgePopedomTF(Session("Name"),"P060503")) OR (JudgePopedomTF(Session("Name"),"P060603"))) then Call ReturnError()
Dim FileID,JSDellObj,FileObj,FileSavaPath,FileName
if Request("FileID") <> "" then
	FileID = Replace(Request("FileID"),"***",",")
	Set JSDellObj = Conn.Execute("select FileSavePath,FileName from FS_SysJS where ID in (" & FileID & ")")
	if JSDellObj.eof then
		Response.Write("<script>alert(""参数传递错误"");window.close();</script>")
		response.end
	end if
else
	Response.Write("<script>alert(""参数传递错误"");window.close();</script>")
	response.end
end if 
%>
<html>
<head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>系统JS删除</title>
</head>
<body leftmargin="0" topmargin="0">
<form action="" name="JSDellForm" method="post">
<table width="100%" border="0" cellspacing="5" cellpadding="0">
  <tr> 
    <td width="7%" height="10">&nbsp;</td>
    <td width="28%" rowspan="3"><div align="center"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
    <td width="59%">&nbsp;</td>
    <td width="6%" height="10">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>您确定要删除此JS?</td>
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
if request.Form("action")="trues" then
	Set FileObj = Server.CreateObject(G_FS_FSO)
	do while Not JSDellObj.Eof
		FileName = JSDellObj("FileName")
		FileSavaPath = JSDellObj("FileSavePath")
		if FileObj.FileExists(Server.MapPath("\")&FileSavaPath&"/"& FileName &".js") = True then
			FileObj.DeleteFile (Server.MapPath("\")&FileSavaPath&"/"& FileName &".js")
		end if
		JSDellObj.MoveNext
	Loop
	Set JSDellObj = Nothing
	Conn.Execute("delete from FS_SysJS where ID in (" & FileID & ")")
	Response.Write("<script>dialogArguments.location.reload();window.close();</script>")
	Response.End
end if
Set JSDellObj = Nothing
%>