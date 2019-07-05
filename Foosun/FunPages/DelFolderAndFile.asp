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
<!--#include file="../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P990400") then Call ReturnError()

Dim DelFolder,Path,DelFile,Action
Path = Request("Path")
DelFolder = Request("DelFolder")
DelFile = Request("DelFile")
Action = Request("Action")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
</head>
<link href="../../CSS/ModeWindow.css" rel="stylesheet">
<body topmargin="0" leftmargin="0" scroll=no>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <form action="" method="post" name="Form">
  <tr> 
    <td height="36" colspan="2"> 
      <div align="center">确定要删除吗？</div></td>
  </tr>
  <tr> 
    <td height="26">
<div align="center">
        <input type="submit" name="Submit" value=" 确 定 ">
          <input name="Action" type="hidden" id="Action" value="Submit">
          <input name="Path" type="hidden" value="<% = Path %>" id="Path">
          <input name="DelFolder" type="hidden" value="<% = DelFolder %>" id="Folder">
          <input name="DelFile" type="hidden" value="<% = DelFile %>" id="FileName">
        </div></td>
    <td height="26">
<div align="center">
        <input type="button" onClick="window.close();" name="Submit2" value=" 取 消 ">
      </div></td>
  </tr>
 </form>
</table>
</body>
</html>
<%
if Action = "Submit" then
	Dim FsoObj,ResponseStr,FileName
	Set FsoObj = Server.CreateObject(G_FS_FSO)
	if Path <> "" then
		if DelFile <> "" then
			Dim DelFileArray,DelFile_i,Temp_False,Temp_FileName
			Temp_False = 0
			Temp_FileName = ""
			DelFileArray = Split(DelFile,"***")
			For DelFile_i = 0 to UBound(DelFileArray)
				If Right(Path,1)="/" then Path=Left(Path,Len(Path)-1) else
				FileName = Server.MapPath(Path & "\" & DelFileArray(DelFile_i))
				if FsoObj.FileExists(FileName) then
					FsoObj.DeleteFile FileName
				else
					Temp_False = Temp_False + 1
					If Temp_FileName = "" then
						Temp_FileName = DelFileArray(DelFile_i)
					Else
						Temp_FileName = Temp_FileName &"|"& DelFileArray(DelFile_i)
					End If
				end if
			Next
			If Temp_False >= 1 then
				ResponseStr = "文件" & Temp_FileName & "删除失败"
			Else
				ResponseStr = ""
			End if
		end if
		if DelFolder <> "" then
			Dim DelFolderArray,DelFolder_i,DelFolder_False,DelFolder_Name,FolderName
			DelFolder_False = 0
			DelFolder_Name = ""
			DelFolderArray = Split(DelFolder,"***")
			For DelFolder_i = 0 to UBound(DelFolderArray)
				if Path = "\" then
					FolderName = Server.MapPath(Path & DelFolderArray(DelFolder_i))
				else
					FolderName = Server.MapPath(Path & "\" & DelFolderArray(DelFolder_i))
				end if
				if FsoObj.FolderExists(FolderName)=true then
					FsoObj.DeleteFolder FolderName
				else
					DelFolder_False = DelFolder_False + 1
					If DelFolder_Name = "" then
						DelFolder_Name = DelFolderArray(DelFolder_i)
					Else
						DelFolder_Name = DelFolder_Name &"|"& DelFolderArray(DelFolder_i)
					End If
				end if
			Next
			If DelFolder_False >= 1 then
				if ResponseStr <> "" then
					ResponseStr = ResponseStr & ";目录" & DelFolder_Name & "删除失败"
				else
					ResponseStr = "目录" & DelFolder_Name & "删除失败"
				end if
			End if
		end if
	else
		ResponseStr = "参数传递错误"
	end if
	Set FsoObj = Nothing
	if ResponseStr <> "" then
		%>
			<script language="JavaScript">alert('<% = ResponseStr %>');dialogArguments.location.reload();window.close();</script>
		<%
	else
		%>
			<script language="JavaScript">dialogArguments.location.reload();window.close();</script>
		<%
	end if
end if
%>