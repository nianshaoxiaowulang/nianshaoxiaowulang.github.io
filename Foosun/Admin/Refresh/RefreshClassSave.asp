<% Option Explicit %>
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
Dim NewsNo,StrClassID
NewsNo = Request("NewsNo")
if NewsNo = "" then
	NewsNo = 0
else
	NewsNo = CInt(NewsNo)
end if
if Not JudgePopedomTF(Session("Name"),"P030300") then Call ReturnError1()
Dim Types,SearchSql,RsSearchObj,NewsTotalNum,PromptInfo,AlreadyRefreshID,AllClassID
AlreadyRefreshID = Request("AlreadyRefreshID")
Types = Request("Types")
StrClassID=Request("ClassID")
if Types = "ClassOne" then
	If Instr(1,StrClassID,",")=0 then
		If Request("IssueSubClass")="IssueSubClass" then
			If Session("AllClassID")="" then
				Session("AllClassID")="'" & StrClassID & "'" & AllChildClassIDList(Request("ClassID"))
			End If
		else
			If Session("AllClassID")="" then
				Session("AllClassID")="'"&StrClassID&"'"
			End If
		End If
	Else
		If Session("AllClassID")="" then
			Session("AllClassID")="'" & replace(replace(StrClassID,",","','")," ","") & "'"
		End If
	End If
	if NewsNo = 0 then
		SearchSql = "Select top 1 * from FS_NewsClass where DelFlag=0 and ClassID in(" & Session("AllClassID") & ") order by ID" 
	else
		SearchSql = "Select top 1 * from FS_NewsClass where  DelFlag=0 and ID>" & AlreadyRefreshID & " and ClassID in(" & Session("AllClassID") & ") order by ID"
	end if
	If session("NewsTotalNum")="" then
		session("NewsTotalNum")= Conn.Execute("Select count(*) from FS_NewsClass where DelFlag=0 and ClassID in(" & Session("AllClassID") & ")")(0)
	End If
elseif Types = "ClassAll" then
	if AlreadyRefreshID = "" then
		SearchSql = "Select top 1 * from FS_NewsClass where DelFlag=0 and IsOutClass=0 order by ID"
	else
		SearchSql = "Select top 1 * from FS_NewsClass where ID>" & AlreadyRefreshID & " and DelFlag=0 and IsOutClass=0 order by ID"
	end if
	If session("NewsTotalNum")="" then
		session("NewsTotalNum")= Conn.Execute("Select count(*) from FS_NewsClass where Delflag=0 and IsOutClass=0")(0)
	End If
else
	SearchSql = ""
	session("NewsTotalNum") = 0
end if
if SearchSql <> "" then
	Set RsSearchObj = Server.CreateObject(G_FS_RS)
	RsSearchObj.Open SearchSql,Conn,1,1
	if RsSearchObj.Eof then
		PromptInfo = "共刷新<b>" & session("NewsTotalNum") & "</b>栏目,刷新成功<font color=red><b>" & Session("RefreshSuccessClass") & "</b></font>个栏目<br>" & Session("RefreshFindTwoLastClass") & "<br> <input name=""imageField"" type=""image"" src=""../../Images/Btn_Back.gif"" width=""75"" height=""21"" border=""0"">"
		Session("RefreshFindTwoLastClass") = ""
		Session("RefreshSuccessClass") = 0
		Session("AllClassID")=""
		Session("NewsTotalNum")=""
		Call PromptFunction
	else
		RefreshClass RsSearchObj
		NewsNo = NewsNo + 1
		AlreadyRefreshID = RsSearchObj("ID")
		Response.Write("<meta http-equiv=""refresh"" content=""0;url=RefreshClassSave.asp?NewsNo=" & NewsNo & "&Types=" & Types & "&AlreadyRefreshID=" & AlreadyRefreshID & "&IssueSubClass="&Request("IssueSubClass")&""">")
		PromptInfo = "共有<font color=red><b>" & session("NewsTotalNum") & "</b></font>个栏目需要刷新<br><br>正在刷新第<font color=red><b>" & NewsNo & "</b></font>个栏目"
		PromptInfo = PromptInfo & "按确定键返回！<br><br><input name=""imageField"" type=""image"" src=""../../Images/Btn_Back.gif"" width=""75"" height=""21"" border=""0"">"
		Call PromptFunction
	end if
	Set RsSearchObj = Nothing
else
	PromptInfo = "共刷新<b>" & session("NewsTotalNum") & "</b>栏目,刷新成功<font color=red><b>" & Session("RefreshSuccessClass") & "</b></font>个栏目<br>" & Session("RefreshFindTwoLastClass") & "<br> <input name=""imageField"" type=""image"" src=""../../Images/Btn_Back.gif"" width=""75"" height=""21"" border=""0"">"
	Session("RefreshFindTwoLastClass") = ""
	Session("RefreshSuccessClass") = 0
	Session("AllClassID")=""
	Session("NewsTotalNum")=""
	Call PromptFunction
end if

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
		FileContent = "模板不存在，请添加模板后再生成！"
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
			FileContent = "模板内容为空"
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
				Session("RefreshFindTwoLastClass") = RefreshClassRsObj("ClassCName") & "栏目捆绑的模板" & RefreshClassRsObj("ClassTemp") & "发现两个终极列表"
			else
				if InStr(Session("RefreshFindTwoLastClass"),RefreshClassRsObj("ClassCName") & "栏目捆绑的模板" & RefreshClassRsObj("ClassTemp") & "发现两个终极列表") = 0 then
					Session("RefreshFindTwoLastClass") = Session("RefreshFindTwoLastClass") & "<br>" & RefreshClassRsObj("ClassCName") & "栏目捆绑的模板" & RefreshClassRsObj("ClassTemp") & "发现两个终极列表"
				end if
			end if
			Exit Function
		end if
	else
		FileContent = Array(FileContent)
	end if
	Session("RefreshSuccessClass") = Session("RefreshSuccessClass") + 1
'	if (Not IsNull(RefreshClassRsObj("DoMain"))) And (RefreshClassRsObj("DoMain") <> "") then
'		CheckFolderExists TempSysRootDir & RefreshClassRsObj("SaveFilePath"),"","","index","0"
'		SaveFilePath = TempSysRootDir & RefreshClassRsObj("SaveFilePath") & "/" & "index"
'	else
		CheckFolderExists TempSysRootDir & RefreshClassRsObj("SaveFilePath"),RefreshClassRsObj("ClassEName"),"","index","0"
		if RefreshClassRsObj("SaveFilePath") = "/" then
			SaveFilePath = TempSysRootDir & RefreshClassRsObj("SaveFilePath") & RefreshClassRsObj("ClassEName") & "/" & "index"
		else
			SaveFilePath = TempSysRootDir & RefreshClassRsObj("SaveFilePath") & "/" & RefreshClassRsObj("ClassEName") & "/" & "index"
		end if
'	end if
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
End Function
Function AllChildClassIDList(ClassID)
	Dim TempRs
	Set TempRs = Conn.Execute("Select ClassID,ChildNum from FS_NewsClass where ParentID = '" & ClassID & "' and DelFlag=0 order by AddTime desc")
	do while Not TempRs.Eof
		AllChildClassIDList = AllChildClassIDList & ",'" & TempRs("ClassID") & "'"
		AllChildClassIDList = AllChildClassIDList & AllChildClassIDList(TempRs("ClassID"))
		TempRs.MoveNext
	loop
	TempRs.Close
	Set TempRs = Nothing
End Function

Sub PromptFunction()
	Set Conn = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>
<link rel="stylesheet" href="../../../CSS/FS_css.css">
<body>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <form method=post action=RefreshClass.asp>
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