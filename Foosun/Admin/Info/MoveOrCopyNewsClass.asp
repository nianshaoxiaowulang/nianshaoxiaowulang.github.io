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
'������ƣ���Ѷ��վ��Ϣ����ϵͳ
'��ǰ�汾��Foosun Content Manager System(FoosunCMS V3.1.0930)
'���¸��£�2005.10
'==============================================================================
'Copyright (C) 2002-2004 Foosun.Net  All rights reserved.
'��ҵע����ϵ��028-85098980-601,��Ŀ������028-85098980-606��609,�ͻ�֧�֣�608
'��Ʒ��ѯQQ��394226379,159410,125114015
'����֧��QQ��315485710,66252421 
'��Ŀ����QQ��415637671��655071
'���򿪷����Ĵ���Ѷ�Ƽ���չ���޹�˾(Foosun Inc.)
'Email:service@Foosun.cn
'MSN��skoolls@hotmail.com
'��̳֧�֣���Ѷ������̳(http://bbs.foosun.net)
'�ٷ���վ��www.Foosun.cn  ��ʾվ�㣺test.cooin.com 
'��վͨϵ��(���ܿ��ٽ�վϵ��)��www.ewebs.cn
'==============================================================================
'��Ѱ汾���ڳ�����ҳ������Ȩ��Ϣ�������ϱ�վLOGO��������
'��Ѷ��˾�����˳���ķ���׷��Ȩ��
'�������2�ο��������뾭����Ѷ��˾������������׷����������
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
		TxtMoveTF = "�ƶ�"
	else
		TxtMoveTF = "����"
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
		TxtObjectClass = "ϵͳ����Ŀ"
	end if
	OperationType = GetParaValue(MoveOrCopyClassPara,"OperationType")
	if OperationType = "Class" then
		if Not JudgePopedomTF(Session("Name"),"P010400") then Call ReturnError()
		TxtOperationType = "��Ŀ"
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
			ShowStr = "ȷ��Ҫ��" & """" & TxtSourceClass & """��Ŀ" & TxtMoveTF & "��""" & TxtObjectClass & """��"
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
		TxtOperationType = "����"
		SourceClass = ""
		SourceNews = Trim(GetParaValue(MoveOrCopyClassPara,"SourceNews"))
		SourceDownLoad = Trim(GetParaValue(MoveOrCopyClassPara,"SourceDownLoad"))
		ShowStr = CheckMoveOrCopyNews(SourceNews,ObjectClass)
		if ShowStr <> "" then
			ShowSubmitTF = false
		else
			ShowSubmitTF = true
			ShowStr = "ȷ��Ҫ" & TxtMoveTF & "��"
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
	ShowStr = "�������ݴ���"
end if

Function CheckMoveOrCopyNews(SourceNewsID,ObjectClassID)
	Dim RsTempObj,TempSourceClassID
	if ObjectClassID = "0" then
		CheckMoveOrCopyNews = "Ŀ�겻���ڣ������ƶ�ʧ��"
		Exit Function
	end if
	'Set RsTempObj = Conn.Execute("Select ClassID from News where NewsID in ('" & Replace(SourceNewsID,"***","','") & "')")
	'if Not RsTempObj.Eof then
		'TempSourceClassID = RsTempobj("ClassID")
	'else
		'CheckMoveOrCopyNews = "���ŵ���Ŀ�����ڣ������ƶ�ʧ��"
		'Exit Function
	'end if
	'Set RsTempObj = Nothing
	'if TempSourceClassID = ObjectClassID then
		'CheckMoveOrCopyNews = "Դ��Ŀ����ͬ�������ƶ�ʧ��"
		'Exit Function
	'end if
End Function

Function CheckMoveOrCopyClass(SourceClassID,ObjectClassID)
	Dim TempClassIDArray,TempLoopVar
	if ObjectClassID <> "0" then
		if InStr(SourceClassID,ObjectClassID) <> 0 then
			CheckMoveOrCopyClass = "Ŀ����Ŀ��Դ��Ŀ��ͬ������ʧ��"
			Exit Function
		end if
	end if
	if SourceClassID = "" then
		CheckMoveOrCopyClass = "Դ��Ŀ�����ڣ�����ʧ��"
		Exit Function
	end if
	if ObjectClassID = "" then
		CheckMoveOrCopyClass = "Ŀ����Ŀ�����ڣ�����ʧ��"
		Exit Function
	end if
	TempClassIDArray = Split(SourceClassID,",")
	for TempLoopVar = LBound(TempClassIDArray) to UBound(TempClassIDArray)
		if JudgeSourceObjectClass(TempClassIDArray(TempLoopVar),ObjectClassID) = true then
			CheckMoveOrCopyClass = "���ܰѸ���Ŀ���������ƶ�������Ŀ������ʧ��"
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
<title>�ƶ����߿���������Ŀ</title>
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
          <input type="submit" name="Submit" value=" ȷ �� ">
          <input name="Result" type="hidden" id="Result" value="Submit">
          <input type="button" name="Submit2" onClick="window.close();" value=" ȡ �� ">
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
