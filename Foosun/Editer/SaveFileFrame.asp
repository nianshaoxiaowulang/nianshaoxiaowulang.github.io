<% Option Explicit %>
<!--#include file="../../Inc/Const.asp" -->
<!--#include file="../../Inc/Cls_DB.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../Inc/Session.asp" -->
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
Dim Result
Result = Request.Form("Result")
if Result = "Submit" then
	Dim Path,FileName,FileContent,EditFile,TempForVar
	Path = Request.Form("Path")
	FileName = Request.Form("FileName")
	For TempForVar = 1 To Request.Form("FileContent").Count
		FileContent = FileContent & Request.Form("FileContent")(TempForVar)
	Next
	'FileContent = Request.Form("FileContent")
	On Error Resume Next
	EditFile = Server.MapPath(Path) & "\" & FileName
	Dim FsoObj,FileObj,FileStreamObj
	Set FsoObj = Server.CreateObject(G_FS_FSO)
	Set FileObj = FsoObj.GetFile(EditFile)
	Set FileStreamObj = FileObj.OpenAsTextStream(2)
	FileStreamObj.Write FileContent
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
