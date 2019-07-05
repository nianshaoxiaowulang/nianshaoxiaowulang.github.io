<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../Refresh/Function.asp" -->
<%
	Dim DBC,Conn
	On Error Resume Next
	Set DBC = New DataBaseClass
	Set Conn = DBC.OpenConnection()
	Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P030705") then Call ReturnError1()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>
<body>
<form action="" method="post" id="SaveFileForm" name="SaveFileForm">
<input name="Path" type="hidden" value="">
<input name="FileName" type="hidden" value="">
<input name="FileContent" type="hidden" value="">
<input name="Result" type="hidden" value="">
</form>
</body>
</html>
<%
Dim Result,TempForVar
Result = Request.Form("Result")
if Result = "Submit" then
	Dim Path,FileName,FileContent,EditFile
	Path = Request.Form("Path")
	FileName = Request.Form("FileName")
	For TempForVar = 1 To Request.Form("FileContent").Count
		FileContent = FileContent & Request.Form("FileContent")(TempForVar)
	Next
	FileContent = "<html>" & Chr(13) & Chr(10) & FileContent & Chr(13) & Chr(10) & "</html>"
	On Error Resume Next
	EditFile = Server.MapPath(Path) & "\" & FileName
	Dim FsoObj,FileObj,FileStreamObj
	Set FsoObj = Server.CreateObject(G_FS_FSO)
	Set FileObj = FsoObj.GetFile(EditFile)
	Set FileStreamObj = FileObj.OpenAsTextStream(2)
	FileStreamObj.Write Replace(FileContent,WebDomain,"")
	if Err.Number <> 0 then
%>
<script language="JavaScript">
alert('<% = "保存失败，请拷贝后，重新打开文件再保存" %>'); //
</script>
<%
	end if
end if
%>
<%
Set FsoObj = Nothing
Set FileObj = Nothing
Set FileStreamObj = Nothing
%>
