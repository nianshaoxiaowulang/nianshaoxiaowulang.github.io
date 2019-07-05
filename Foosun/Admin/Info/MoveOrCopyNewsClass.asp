<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../Inc/Cls_Info.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
'==============================================================================
'软件名称：风讯网站信息管理系统
'当前版本：Foosun Content Manager System(FoosunCMS V3.1.0930)
'最新更新：2005.10
'==============================================================================
'Copyright (C) 2002-2004 Foosun.Net  All rights reserved.
'商业注册联系：028-85098980-601,项目开发：028-85098980-606、609,客户支持：608
'产品咨询QQ：394226379,159410,125114015
'技术支持QQ：315485710,66252421 
'项目开发QQ：415637671，655071
'程序开发：四川风讯科技发展有限公司(Foosun Inc.)
'Email:service@Foosun.cn
'MSN：skoolls@hotmail.com
'论坛支持：风讯在线论坛(http://bbs.foosun.net)
'官方网站：www.Foosun.cn  演示站点：test.cooin.com 
'网站通系列(智能快速建站系列)：www.ewebs.cn
'==============================================================================
'免费版本请在程序首页保留版权信息，并做上本站LOGO友情连接
'风讯公司保留此程序的法律追究权利
'如需进行2次开发，必须经过风讯公司书面允许。否则将追究法律责任
'==============================================================================

%>
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<!--#include file="../../../Inc/Session.asp" -->
<%
if Not ((JudgePopedomTF(Session("Name"),"P010400")) OR (JudgePopedomTF(Session("Name"),"P010503"))) then Call ReturnError()
Dim MoveOrCopyClassPara,ShowStr,i,LoopVar,IDStr
Dim ShowSubmitTF,Result
Result = Request.Form("Result")
ShowSubmitTF = true
MoveOrCopyClassPara = Request("MoveOrCopyClassPara")
if MoveOrCopyClassPara <> "" then
	Dim OperationType,MoveTF,SourceClass,SourceNews,ObjectClass,RsTempObj
	Dim TxtOperationType,TxtMoveTF,TxtSourceClass,TxtSourceNews,TxtObjectClass
	MoveTF = GetParaValue(MoveOrCopyClassPara,"MoveTF")
	if MoveTF = "true" then
		TxtMoveTF = "移动"
	else
		TxtMoveTF = "拷贝"
	end if
	ObjectClass = GetParaValue(MoveOrCopyClassPara,"ObjectClass")
	if ObjectClass <> "0" then
		Set RsTempObj = Conn.Execute("Select ClassCName from FS_NewsClass where ClassID='" & ObjectClass & "'")
		if Not RsTempObj.Eof then
			TxtObjectClass = RsTempObj("ClassCName")
		else
			TxtObjectClass = ""
		end if
		Set RsTempObj = Nothing
	else
		TxtObjectClass = "系统根栏目"
	end if
	OperationType = GetParaValue(MoveOrCopyClassPara,"OperationType")
	if OperationType = "Class" then
		if Not JudgePopedomTF(Session("Name"),"P010400") then Call ReturnError()
		TxtOperationType = "栏目"
		SourceClass = GetParaValue(MoveOrCopyClassPara,"SourceClass")
		IDStr = Replace(SourceClass,",","','")
		Set RsTempObj = Conn.Execute("Select ClassCName from FS_NewsClass where ClassID in ('" & IDStr & "')")
		if Not RsTempObj.Eof then
			do while Not RsTempObj.Eof
				if TxtSourceClass = "" then
					TxtSourceClass = RsTempObj("ClassCName")
				else
					TxtSourceClass = TxtSourceClass & "|" & RsTempObj("ClassCName")
				end if
				RsTempObj.MoveNext
			Loop
		else
			TxtSourceClass = ""
		end if
		Set RsTempObj = Nothing
		SourceNews = ""
		ShowStr = CheckMoveOrCopyClass(SourceClass,ObjectClass)
		if ShowStr <> "" then
			ShowSubmitTF = false
		else
			ShowSubmitTF = true
			ShowStr = "确定要把" & """" & TxtSourceClass & """栏目" & TxtMoveTF & "到""" & TxtObjectClass & """吗？"
		end if
		if Result = "Submit" then
			Dim CClass
			Set CClass = New InfoClass
			if MoveTF = "true" then
				CClass.MoveClass SourceClass,ObjectClass
			else
				CClass.CopyClass SourceClass,ObjectClass 
			end if
			Set CClass = Nothing
			%>
				<script language="JavaScript">
				dialogArguments.top.GetNavFoldersObject().location='../Menu_Folders.asp?Action=ContentTree&OpenClassIDList=<% = ParentClassIDList(ObjectClass) & ObjectClass & "," & SourceClass %>';		
				window.close();
				</script>
			<%
		end if
	else
		if Not JudgePopedomTF(Session("Name"),"P010503") then Call ReturnError()
		Dim NewsPromptStr,DownLoadPromptStr,SourceDownLoad
		TxtOperationType = "内容"
		SourceClass = ""
		SourceNews = Trim(GetParaValue(MoveOrCopyClassPara,"SourceNews"))
		SourceDownLoad = Trim(GetParaValue(MoveOrCopyClassPara,"SourceDownLoad"))
		ShowStr = CheckMoveOrCopyNews(SourceNews,ObjectClass)
		if ShowStr <> "" then
			ShowSubmitTF = false
		else
			ShowSubmitTF = true
			ShowStr = "确定要" & TxtMoveTF & "吗？"
		end if
		if Result = "Submit" then
			Dim NClass,SourceNewsArray,SourceDownLoadArray
			SourceNewsArray = Split(SourceNews,"***")
			SourceDownLoadArray = Split(SourceDownLoad,"***")
			Set NClass = New InfoClass
			if MoveTF = "true" then
				NClass.MoveNews SourceNewsArray,ObjectClass
				NClass.MoveDownLoad SourceDownLoadArray,ObjectClass
			else
				NClass.CopyNews SourceNewsArray,ObjectClass 
				NClass.CopyDownLoad SourceDownLoadArray,ObjectClass
			end if
			Set NClass = Nothing
			%>
			<script language="JavaScript">
			window.close();
			</script>
			<%
		end if
	end if
else
	ShowSubmitTF = false
	ShowStr = "参数传递错误"
end if

Function CheckMoveOrCopyNews(SourceNewsID,ObjectClassID)
	Dim RsTempObj,TempSourceClassID
	if ObjectClassID = "0" then
		CheckMoveOrCopyNews = "目标不存在，新闻移动失败"
		Exit Function
	end if
	'Set RsTempObj = Conn.Execute("Select ClassID from News where NewsID in ('" & Replace(SourceNewsID,"***","','") & "')")
	'if Not RsTempObj.Eof then
		'TempSourceClassID = RsTempobj("ClassID")
	'else
		'CheckMoveOrCopyNews = "新闻的栏目不存在，新闻移动失败"
		'Exit Function
	'end if
	'Set RsTempObj = Nothing
	'if TempSourceClassID = ObjectClassID then
		'CheckMoveOrCopyNews = "源和目的相同，新闻移动失败"
		'Exit Function
	'end if
End Function

Function CheckMoveOrCopyClass(SourceClassID,ObjectClassID)
	Dim TempClassIDArray,TempLoopVar
	if ObjectClassID <> "0" then
		if InStr(SourceClassID,ObjectClassID) <> 0 then
			CheckMoveOrCopyClass = "目标栏目和源栏目相同，操作失败"
			Exit Function
		end if
	end if
	if SourceClassID = "" then
		CheckMoveOrCopyClass = "源栏目不存在，操作失败"
		Exit Function
	end if
	if ObjectClassID = "" then
		CheckMoveOrCopyClass = "目标栏目不存在，操作失败"
		Exit Function
	end if
	TempClassIDArray = Split(SourceClassID,",")
	for TempLoopVar = LBound(TempClassIDArray) to UBound(TempClassIDArray)
		if JudgeSourceObjectClass(TempClassIDArray(TempLoopVar),ObjectClassID) = true then
			CheckMoveOrCopyClass = "不能把父栏目拷贝或者移动到子栏目，操作失败"
			Exit Function
		end if
	Next
End Function

Private Function JudgeSourceObjectClass(SourceClassID,ObjectClassID)
	Dim TempSql,RsTempObj,Temp
	TempSql = "Select ClassID from FS_NewsClass where ParentID ='" & SourceClassID & "'"
	Set RsTempObj = Conn.Execute(TempSql)
	do while Not RsTempObj.Eof
		if RsTempObj("ClassID") = ObjectClassID then
			JudgeSourceObjectClass = True
			Exit do
		end if
		JudgeSourceObjectClass = JudgeSourceObjectClass(RsTempObj("ClassID"),ObjectClassID)
		if JudgeSourceObjectClass = True then Exit do
		RsTempObj.MoveNext
	loop
	RsTempObj.Close
	Set RsTempObj = Nothing
End Function

Function GetParaValue(ParaStr,ParaName)
	Dim BeginIndex,EndIndex
	BeginIndex = InStr(ParaStr,ParaName)+Len(ParaName)+1
	EndIndex = InStr(BeginIndex,ParaStr,",")
	GetParaValue = Mid(ParaStr,BeginIndex,EndIndex-BeginIndex)
End Function

Function ParentClassIDList(ClassID)
	Dim TempRs,TempStr
	Set TempRs = Conn.Execute("Select ParentID from FS_NewsClass where ClassID = '" & ClassID & "'")
	if Not TempRs.Eof then
		if TempRs("ParentID") <> "0" then
			ParentClassIDList =  TempRs("ParentID") & "," & ParentClassIDList
			ParentClassIDList = ParentClassIDList & ParentClassIDList(TempRs("ParentID"))
		end if
	end if
	TempRs.Close
	Set TempRs = Nothing
End Function
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>移动或者拷贝新闻栏目</title>
</head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet">
<body topmargin="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <form name="OperateForm" action="" method="post">
    <tr> 
      <td height="10" colspan="2"></td>
    </tr>
    <tr>
      <td width="28%" rowspan="3"><div align="center"><strong><font size="2"><img src="../../Images/Question.gif" width="39" height="37"></font></strong></div></td>
      <td height="5"></td>
    </tr>
    <tr>
      <td width="72%"><% = ShowStr %></td>
    </tr>
    <tr>
      <td width="72%" height="10"></td>
    </tr>
    <tr> 
<%
if ShowSubmitTF = true then
%>
      <td colspan="2"><div align="center"> 
          <input type="submit" name="Submit" value=" 确 定 ">
          <input name="Result" type="hidden" id="Result" value="Submit">
          <input type="button" name="Submit2" onClick="window.close();" value=" 取 消 ">
        </div></td>
<%
end if
%>
    </tr>
  </form>
</table>
</body>
</html>
<%
Set Conn = Nothing
%>
