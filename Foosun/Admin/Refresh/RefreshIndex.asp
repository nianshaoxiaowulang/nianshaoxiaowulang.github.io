<% Option Explicit %>
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="RefreshFunction.asp" -->
<!--#include file="SelectFunction.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="Function.asp" -->
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P030100") then Call ReturnError1()
Dim SaveFilePath,FSOObj,FSOObj1,PromptInfo,FileStreamObj,FileContent,FileObj,TempletFileName,RefreshTempFileName
PromptInfo = ""
if SysRootDir = "" then
	TempletFileName = Server.MapPath("/" & TempletDir) & "\Index.htm"
	SaveFilePath = "/index."&confimsn("IndexExtName")&""
else
	TempletFileName = Server.MapPath("/" & SysRootDir & "/" & TempletDir) & "\Index.htm"
	SaveFilePath = "/" & SysRootDir & "/index."&confimsn("IndexExtName")&""
end if
Set FSOObj = Server.CreateObject(G_FS_FSO)
if FSOObj.FileExists(TempletFileName) = False then
	PromptInfo = "��ҳģ��Index.htm�����ڣ��������ҳģ��������ɣ�"
	Call PromptFunction
else
	'On Error Resume Next
	SetRefreshValue "Index",""
	GetAvailableDoMain
	Set FileObj = FSOObj.GetFile(TempletFileName)
	Set FileStreamObj = FileObj.OpenAsTextStream(1)
	if Not FileStreamObj.AtEndOfStream then
		FileContent = FileStreamObj.ReadAll
		FileContent = ReplaceAllServerFlag(ReplaceAllLable(FileContent))
	else
		FileContent = "ģ������Ϊ��"
	end if
	Set FileStreamObj = Nothing
	Select Case AvailableRefreshType
		Case 0
			FSOSaveFile FileContent,SaveFilePath
		Case 1
			SaveFile FileContent,SaveFilePath
		Case Else
			FSOSaveFile FileContent,SaveFilePath
	End Select
	PromptInfo = "���ɳɹ�"
	Call PromptFunction
end if
Set FSOObj = Nothing
Sub PromptFunction()
	Set Conn = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>������ҳ���ɹ���</title>
</head>
<link rel="stylesheet" href="../../../CSS/FS_css.css">
<body topmargin="2" leftmargin="2" oncontextmenu="return false;">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="28" class="ButtonListLeft">
<div align="center"><strong>������ҳ���ɹ���</strong></div></td>
</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td><div align="center"><font color="#FF0000">������Ϣ: <a href="<%=AvailableDoMain%>/index.<%=confimsn("IndexExtName")%>" target="_blank">�����ҳ</a></font></div></td>
  </tr>
  <tr> 
    <td><div align="center"> 
        <% = PromptInfo %>
      </div></td>
  </tr>
</table>
</body>
</html>
<%
End Sub
%>
