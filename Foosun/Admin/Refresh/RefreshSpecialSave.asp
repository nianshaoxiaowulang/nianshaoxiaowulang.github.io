<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="Function.asp" -->
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
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P030500") then Call ReturnError1()
Dim Types,SearchSql,SpecialNo,RsSearchObj,SpecialTotalNum,PromptInfo,SpecialID,RefreshOne
SpecialID=Request.Form("SpecialID")
SpecialNo = Request("SpecialNo")
if SpecialNo = "" then
	SpecialNo = 1
else
	SpecialNo = CInt(SpecialNo)
end if
Types = Request("Types")
if Types = "SpecialOne" then
	SearchSql = "Select * from FS_Special where SpecialID='" & SpecialID & "'"
	Types="RefreshOver"
elseif Types = "SpecialAll" then
	SearchSql = "Select * from FS_Special"
else
	SearchSql = ""
end if
if SearchSql <> "" then
	Set RsSearchObj = Server.CreateObject(G_FS_RS)
	RsSearchObj.Open SearchSql,Conn,1,1
	SpecialTotalNum = RsSearchObj.RecordCount
	if RsSearchObj.Eof and types<>"RefreshOver" then
		PromptInfo = "û��Ҫˢ�µ�ר��&nbsp;&nbsp;<font color=""red""><a href=""RefreshSpecial.asp"">����</a></font>"
		Set RsSearchObj = Nothing
		Call PromptFunction
	else
		RsSearchObj.Move SpecialNo - 1
		if Not RsSearchObj.Eof then
			RefreshSpecial RsSearchObj
			SpecialNo = SpecialNo + 1
			Response.Write("<meta http-equiv=""refresh"" content=""0;url=RefreshSpecialSave.asp?SpecialNo=" & SpecialNo &"&Types="&types& "&SpecialID="&SpecialID&""">")
			PromptInfo = "����<font color=red><b>" & SpecialTotalNum & "</b></font>��ר����Ҫˢ��<br><br>����ˢ�µ�<font color=red><b>" & SpecialNo - 1 & "</b></font>��ר��"
			PromptInfo = PromptInfo & "��ȷ�������أ�<br><br><input name=""imageField"" type=""image"" src=""../../Images/Btn_Back.gif"" width=""75"" height=""21"" border=""0"">"
		else
			PromptInfo = "��ˢ��<b>" & SpecialTotalNum & "</b>ר��,ˢ�³ɹ�<font color=red><b>" & Session("RefreshSuccessClass") & "</b></font>��ר��<br>" & Session("RefreshFindTwoLastClass") & "<br> <input name=""imageField"" type=""image"" src=""../../Images/Btn_Back.gif"" width=""75"" height=""21"" border=""0"">"
			Session("RefreshFindTwoLastClass") = ""
			Session("RefreshSuccessClass") = 0
		end if
		Set RsSearchObj = Nothing
		Call PromptFunction
	end if
	Set RsSearchObj = Nothing
else
	PromptInfo = "��ˢ��<b>" & SpecialNo-1 & "</b>ר��,ˢ�³ɹ�<font color=red><b>" & Session("RefreshSuccessClass") & "</b></font>��ר��<br>" & Session("RefreshFindTwoLastClass") & "<br> <input name=""imageField"" type=""image"" src=""../../Images/Btn_Back.gif"" width=""75"" height=""21"" border=""0"">"
	Session("RefreshFindTwoLastClass") = ""
	Session("RefreshSuccessClass") = 0
	Call PromptFunction
end if

Function RefreshSpecial(RefreshSpecialRsObj)
	Dim FSOObj,FileObj,FileStreamObj,TempletFileName,SaveFilePath,FileContent
	Dim TempArray,LoopVar,TempReturnValue,ReturnLoopVar
	Dim TempSysRootDir
	SetRefreshValue "Special",RefreshSpecialRsObj("SpecialID")
	GetAvailableDoMain
	if SysRootDir = "" then
		TempSysRootDir = ""
	else
		TempSysRootDir = "/" & SysRootDir
	end if
	TempletFileName = Server.MapPath(TempSysRootDir & RefreshSpecialRsObj("Templet"))
	Set FSOObj = Server.CreateObject(G_FS_FSO)
	if FSOObj.FileExists(TempletFileName) = False then
		FileContent = "ģ�岻���ڣ������ģ��������ɣ�"
	else
		Set FileObj = FSOObj.GetFile(TempletFileName)
		Set FileStreamObj = FileObj.OpenAsTextStream(1)
		if Not FileStreamObj.AtEndOfStream then
			FileContent = FileStreamObj.ReadAll
			FileContent = ReplaceAllServerFlag(ReplaceAllLable(FileContent))
		else
			FileContent = "ģ������Ϊ��"
		end if
	end if
	Set FileStreamObj = Nothing
	Set FileObj = Nothing
	Set FSOObj = Nothing
	CheckFolderExists TempSysRootDir & RefreshSpecialRsObj("SaveFilePath"),RefreshSpecialRsObj("EName"),"","index","0"
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
				Session("RefreshFindTwoLastClass") = RefreshSpecialRsObj("CName") & "ר�������ģ��" & RefreshSpecialRsObj("Templet") & "���������ռ��б�"
			else
				if InStr(Session("RefreshFindTwoLastClass"),RefreshSpecialRsObj("CName") & "ר�������ģ��" & RefreshSpecialRsObj("Templet") & "���������ռ��б�") = 0 then
					Session("RefreshFindTwoLastClass") = Session("RefreshFindTwoLastClass") & "<br>" & RefreshSpecialRsObj("CName") & "ר�������ģ��" & RefreshSpecialRsObj("Templet") & "���������ռ��б�"
				end if
			end if
			Exit Function
		end if
	else
		FileContent = Array(FileContent)
	end if
	Session("RefreshSuccessClass") = Session("RefreshSuccessClass") + 1
	if RefreshSpecialRsObj("SaveFilePath") = "/" then
		SaveFilePath = TempSysRootDir & RefreshSpecialRsObj("SaveFilePath") & RefreshSpecialRsObj("EName") & "/" & "index." & RefreshSpecialRsObj("FileExtName")
	else
		SaveFilePath = TempSysRootDir & RefreshSpecialRsObj("SaveFilePath") & "/" & RefreshSpecialRsObj("EName") & "/" & "index." & RefreshSpecialRsObj("FileExtName")
	end if
	
	for LoopVar = LBound(FileContent) to UBound(FileContent)
		if LoopVar = 0 then
			if RefreshSpecialRsObj("SaveFilePath") = "/" then
				SaveFilePath = TempSysRootDir & RefreshSpecialRsObj("SaveFilePath") & RefreshSpecialRsObj("EName") & "/" & "index." & RefreshSpecialRsObj("FileExtName")
			else
				SaveFilePath = TempSysRootDir & RefreshSpecialRsObj("SaveFilePath") & "/" & RefreshSpecialRsObj("EName") & "/" & "index." & RefreshSpecialRsObj("FileExtName")
			end if
		else
			if RefreshSpecialRsObj("SaveFilePath") = "/" then
				SaveFilePath = TempSysRootDir & RefreshSpecialRsObj("SaveFilePath") & RefreshSpecialRsObj("EName") & "/" & "index_" & LoopVar + 1 & "." & RefreshSpecialRsObj("FileExtName")
			else
				SaveFilePath = TempSysRootDir & RefreshSpecialRsObj("SaveFilePath") & "/" & RefreshSpecialRsObj("EName") & "/" & "index_" & LoopVar + 1 & "." & RefreshSpecialRsObj("FileExtName")
			end if
		end if
		Select Case AvailableRefreshType
			Case 0
				FSOSaveFile FileContent(LoopVar),SaveFilePath
			Case 1
				SaveFile FileContent(LoopVar),SaveFilePath
			Case Else
				FSOSaveFile FileContent(LoopVar),SaveFilePath
		End Select
	Next
End Function
Sub PromptFunction()
	Set Conn = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����ר��</title>
</head>
<link rel="stylesheet" href="../../../CSS/FS_css.css">
<body>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <form method=post action=RefreshSpecial.asp>
    <tr> 
      <td><div align="center"> 
          <% = PromptInfo %>
        </div></td>
    </tr>
  </form>
</table>
</body>
</html>
<%
End Sub
%>