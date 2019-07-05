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
Dim StrCLassID,RsRefreshClass,PromptInfo
StrCLassID=Request.QueryString("ClassID")

If IsNumeric(StrCLassID) then
	If Request.QueryString("AutoClass")="1" then
		if JudgePopedomTF(Session("Name"),"P030300") then'������ĿȨ��
			Set RsRefreshClass=Server.CreateObject(G_FS_RS) 
			RsRefreshClass=Conn.Execute("Select * from FS_NewsClass Where ClassID='"&StrCLassID&"'")
			RefreshClass RsRefreshClass
		Else
			PromptInfo=PromptInfo&"��û��������Ŀ��Ȩ�ޣ�<br>"
		End If
	Else
		PromptInfo=PromptInfo&"��û�������Զ�������Ŀ��<br>"
	End If
	If Request.QueryString("AutoIndex")="1" then
		If JudgePopedomTF(Session("Name"),"P030100") then '������ҳȨ��
			ReFreshIndex
		Else
			PromptInfo=PromptInfo & "��û��������ҳ��Ȩ�ޣ�<br>"
		End If
	Else
		PromptInfo=PromptInfo&"��û�������Զ�������ҳ��<br>"
	End If
	PromptFunction
Else
	Response.end
End If


Function ReFreshIndex
	Dim SaveFilePath,FSOObj,FSOObj1,FileStreamObj,FileContent,FileObj,TempletFileName,RefreshTempFileName
	if SysRootDir = "" then
		TempletFileName = Server.MapPath("/" & TempletDir) & "\Index.htm"
		SaveFilePath = "/index."&confimsn("IndexExtName")&""
	else
		TempletFileName = Server.MapPath("/" & SysRootDir & "/" & TempletDir) & "\Index.htm"
		SaveFilePath = "/" & SysRootDir & "/index."&confimsn("IndexExtName")&""
	end if
	Set FSOObj = Server.CreateObject(G_FS_FSO)
	if FSOObj.FileExists(TempletFileName) = False then
		PromptInfo = PromptInfo&"��ҳģ��Index.htm�����ڣ��������ҳģ��������ɣ�"
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
		PromptInfo = PromptInfo &"��ҳ���ɳɹ���"
	end if
	Set FSOObj = Nothing
End Function
'=============================
'������Ŀ
Function RefreshClass(RefreshClassRsObj)
	Dim FSOObj,FileObj,FileStreamObj,TempletFileName,SaveFilePath,FileContent,TempSysRootDir
	Dim TempArray,LoopVar,TempReturnValue,ReturnLoopVar
	SetRefreshValue "Class",RefreshClassRsObj("ClassID")
	GetAvailableDoMain
	if SysRootDir = "" then
		TempSysRootDir = ""
	else
		TempSysRootDir = "/" & SysRootDir
	end if
	TempletFileName = Server.MapPath(TempSysRootDir & RefreshClassRsObj("ClassTemp"))
	Set FSOObj = Server.CreateObject(G_FS_FSO)
	if FSOObj.FileExists(TempletFileName) = False then
		FileContent = "ģ�岻���ڣ������ģ��������ɣ�"
	else
		Set FileObj = FSOObj.GetFile(TempletFileName)
		Set FileStreamObj = FileObj.OpenAsTextStream(1)
		if Not FileStreamObj.AtEndOfStream then
			FileContent = FileStreamObj.ReadAll
			if (RefreshClassRsObj("BrowPop") <> 0) And (RefreshClassRsObj("FileExtName") = "asp") then
				FileContent = GetPopStr(RefreshClassRsObj("BrowPop"),RefreshClassRsObj("SaveFilePath")) & FileContent
			end if
			FileContent = ReplaceAllServerFlag(ReplaceAllLable(FileContent))
		else
			FileContent = "ģ������Ϊ��"
		end if
	end if
	Set FileStreamObj = Nothing
	Set FileObj = Nothing
	Set FSOObj = Nothing
	if NotReplaceLableArray <> "" then
		TempArray = Split(NotReplaceLableArray,"$$$")
		if UBound(TempArray) = 0 then
			if TempArray(0) <> "" then
				TempReturnValue = GetLableContent(TempArray(0))
				for ReturnLoopVar = LBound(TempReturnValue) to UBound(TempReturnValue)
					TempReturnValue(ReturnLoopVar) = Replace(FileContent,Split(NotReplaceLableOldArray,"$$$")(0),TempReturnValue(ReturnLoopVar))
				Next
				FileContent = TempReturnValue
			else
				FileContent = Array(FileContent)
			end if
		else
			if Session("RefreshFindTwoLastClass") = "" then
				Session("RefreshFindTwoLastClass") = RefreshClassRsObj("ClassCName") & "��Ŀ�����ģ��" & RefreshClassRsObj("ClassTemp") & "���������ռ��б�"
			else
				if InStr(Session("RefreshFindTwoLastClass"),RefreshClassRsObj("ClassCName") & "��Ŀ�����ģ��" & RefreshClassRsObj("ClassTemp") & "���������ռ��б�") = 0 then
					Session("RefreshFindTwoLastClass") = Session("RefreshFindTwoLastClass") & "<br>" & RefreshClassRsObj("ClassCName") & "��Ŀ�����ģ��" & RefreshClassRsObj("ClassTemp") & "���������ռ��б�"
				end if
			end if
			Exit Function
		end if
	else
		FileContent = Array(FileContent)
	end if
	Session("RefreshSuccessClass") = Session("RefreshSuccessClass") + 1
	CheckFolderExists TempSysRootDir & RefreshClassRsObj("SaveFilePath"),RefreshClassRsObj("ClassEName"),"","index","0"
	if RefreshClassRsObj("SaveFilePath") = "/" then
		SaveFilePath = TempSysRootDir & RefreshClassRsObj("SaveFilePath") & RefreshClassRsObj("ClassEName") & "/" & "index"
	else
		SaveFilePath = TempSysRootDir & RefreshClassRsObj("SaveFilePath") & "/" & RefreshClassRsObj("ClassEName") & "/" & "index"
	end if
	Dim EndSaveFilePath
	for LoopVar = LBound(FileContent) to UBound(FileContent)
		if LoopVar = 0 then
			EndSaveFilePath =  SaveFilePath & "." & RefreshClassRsObj("FileExtName")
		else
			EndSaveFilePath = SaveFilePath & "_" & LoopVar + 1 & "." & RefreshClassRsObj("FileExtName")
		end if
		Select Case AvailableRefreshType
			Case 0
				FSOSaveFile FileContent(LoopVar),EndSaveFilePath
			Case 1
				SaveFile FileContent(LoopVar),EndSaveFilePath
			Case Else
				FSOSaveFile FileContent(LoopVar),EndSaveFilePath
		End Select
	Next
	PromptInfo=PromptInfo&"��Ŀ���ɳɹ���<br>"
End Function
'===============================


Sub PromptFunction()
	Set Conn = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>���ɹ���</title>
</head>
<link rel="stylesheet" href="../../../CSS/FS_css.css">
<body topmargin="2" leftmargin="2" oncontextmenu="return false;">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="28" class="ButtonListLeft">
<div align="center"><strong>���ɹ���</strong></div></td>
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
