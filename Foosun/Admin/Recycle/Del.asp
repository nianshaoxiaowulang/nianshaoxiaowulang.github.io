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
if Not JudgePopedomTF(Session("Name"),"P070102") then Call ReturnError()
Dim OperateType,NewsID,ClassID
Dim Sql,RsTempObj,PromptInfo
OperateType = Request("OperateType")
NewsID = Request("NewsID")
ClassID=Request("ClassID")
Dim RecSysRootDir
if SysRootDir = "" then
	RecSysRootDir = ""
else
	RecSysRootDir = "/" & SysRootDir
end if

Dim Result
Result = Request("Result")
if Result = "Submit" then
	If ClassID<>"" then
		DelClass ClassID
	End If
	if NewsID <> "" then
		DelNews NewsID
	End if
end if
Set RsTempObj = Nothing
Set Conn = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ޱ����ĵ�</title>
</head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet">
<body scrolling=no>
<table width="100%" border="0" cellspacing="5" cellpadding="0">
  <tr> 
    <td width="30%" align="center"><img src="../../Images/Question.gif" width="39" height="37"> 
    </td>
    <td width="70%" align="center"><div align="left">ȷ��Ҫɾ����?
    </div></td>
  </tr>
  <tr> 
    <td colspan="2"><div align="center"> 
        <input onClick="SubmitFun();" type="button" name="Submit" value=" ȷ �� ">
        <input type="button" onClick="window.close();" name="Submit2" value=" ȡ �� ">
      </div></td>
  </tr>
</table>
</body>
</html>
<script language="JavaScript">
function SubmitFun()
{
	window.location='?NewsID=<% = NewsID%>&ClassID=<% = ClassID %>&Result=Submit';
}
</script>
<%
Function DelClass(DelClassID)
	Dim AllClassID,Sql,i,DelClassIDList
	AllClassID=""
    DelClassIDList=split(DelClassID,"***")
	For i = 0 to ubound(DelClassIDList)
		AllClassID = "'" & DelClassIDList(i) & "'" & ChildClassIDList(DelClassIDList(i))
		On Error Resume Next
		Sql = "Delete from FS_News where ClassID in (" & AllClassID & ")"
		Conn.Execute(Sql)
		if Err.Number <> 0 then Alert "ɾ����Ŀ�µ�����ʧ��"
		Sql = "Delete from FS_Contribution where ClassID in (" & AllClassID & ")"
		Conn.Execute(Sql)
		if Err.Number <> 0 then Alert "ɾ����Ŀ�µ�Ͷ��ʧ��"
		Sql = "Delete from FS_DownLoad where ClassID in (" & AllClassID & ")"
		Conn.Execute(Sql)
		if Err.Number <> 0 then Alert "ɾ����Ŀ�µ�����ʧ��"
		Conn.Execute("Delete from FS_SysJs where ClassID in ("& AllClassID &")")
		If Err.Number <> 0 then Alert "ɾ����Ŀ��ϵͳJSʧ��"
		Conn.Execute("Delete from FS_FreeJsFile where ClassID in (" & AllClassID & ")")
		If Err.Number <> 0 then Alert "ɾ����Ŀ���������JSʧ��"
		'-------------ɾ�������ļ�-----------
		Dim MyFile,DDelClassObj
		Set MyFile=Server.CreateObject(G_FS_FSO)
		Set DDelClassObj = Conn.Execute("Select SaveFilePath,ClassEName from FS_NewsClass where ClassID in (" & AllClassID & ")")
		Do while Not DDelClassObj.eof
			If MyFile.FolderExists(Server.Mappath(RecSysRootDir&DDelClassObj("SaveFilePath")&"/"&DDelClassObj("ClassEName"))) then
				MyFile.DeleteFolder(Server.Mappath(RecSysRootDir&DDelClassObj("SaveFilePath")&"/"&DDelClassObj("ClassEName")))
			End if
			DDelClassObj.MoveNext
		Loop
		DDelClassObj.Close
		Set DDelClassObj = Nothing
		Set MyFile = Nothing
		'------------------------------------
		Sql = "Delete from FS_NewsClass where ClassID in (" & AllClassID & ")"
		Conn.Execute(Sql)
		AllClassID=""
	Next
	if Err.Number = 0 then
		Alert "ɾ���ɹ�"
	else
		Alert "ɾ��ʧ��"
	end if
End Function
Function ChildClassIDList(ClassID)
	Dim TempRs,TempStr
	Set TempRs = Conn.Execute("Select ClassID from FS_NewsClass where ParentID = '" & ClassID & "'")
	do while Not TempRs.Eof
		ChildClassIDList = ChildClassIDList & ",'" & TempRs("ClassID") & "'"
		ChildClassIDList = ChildClassIDList & ChildClassIDList(TempRs("ClassID"))
		TempRs.MoveNext
	loop
	TempRs.Close
	Set TempRs = Nothing
End Function
Function DelNews(DelNewsID)
	Dim Sql,MyFile,i,DelNewsIDList
    DelNewsIDList=split(DelNewsID,"***")
	Set MyFile=Server.CreateObject(G_FS_FSO)
	For i = 0 to ubound(DelNewsIDList)
		Sql = "Delete from FS_News where NewsID='" & DelNewsIDList(i) & "'"
		On Error Resume Next
		Conn.Execute("Delete from FS_FreeJsFile where FileName=(Select FileName from FS_News where NewsID='" & DelNewsIDList(i) & "')")
		
		'------------------------ɾ�����������ļ�-------------------
		Dim DelNewsClassFileObj,DelNewsFileObj
		Set DelNewsFileObj = Conn.Execute("Select FileName,FileExtName,ClassID from FS_News where NewsID='"&DelNewsIDList(i)&"'")
		If Not DelNewsFileObj.eof then
			Set DelNewsClassFileObj = Conn.execute("Select ClassEName,SaveFilePath from FS_NewsClass where ClassID='"&DelNewsFileObj("ClassID")&"'")
			If Not DelNewsFileObj.eof then
				If MyFile.FileExists(Server.Mappath(RecSysRootDir&DelNewsClassFileObj("SaveFilePath")&"/"&DelNewsClassFileObj("ClassEName"))&"/"&DelNewsFileObj("FileName")&"."&DelNewsFileObj("FileExtName")) then
				   MyFile.DeleteFile(Server.Mappath(RecSysRootDir&DelNewsClassFileObj("SaveFilePath")&"/"&DelNewsClassFileObj("ClassEName"))&"/"&DelNewsFileObj("FileName")&"."&DelNewsFileObj("FileExtName"))
				End if
			End If
		End If
		
		Conn.Execute(Sql)
	Next
	if Err.Number = 0 then
		%>
		<script language="JavaScript">
		window.close();
		</script>
		<%
	else
		Alert "ɾ��ʧ��"
	end if
	Set MyFile = Nothing
End Function
Function Alert(InfoStr)
%>
<script language="JavaScript">
alert('<% = InfoStr %>');
window.close();
</script>
<%
End Function
%>