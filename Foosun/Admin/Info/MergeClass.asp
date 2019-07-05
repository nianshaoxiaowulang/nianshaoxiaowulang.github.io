<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../Inc/Cls_Info.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../refresh/Function.asp" -->
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
if Not (JudgePopedomTF(Session("Name"),"P010512")) then Call ReturnError()
if Not JudgePopedomTF(Session("Name"),"" & Request("SourceClass") & "") then Call ReturnError()
if Not JudgePopedomTF(Session("Name"),"" & Request("ObjectClass") & "") then Call ReturnError()
Dim ShowSubmitTF,SourceClass,ObjectClass,Result,ShowStr,RsSoueceObj,RsObjectObj,AllowOperation,AllClassID
SourceClass = Request("SourceClass")
ObjectClass = Request("ObjectClass")
AllClassID = "'" & SourceClass & "'" & ChildClassIDList(SourceClass)
If SourceClass<>ObjectClass then 
	Result = Request("Result")
	ShowSubmitTF = True
	AllowOperation = True
	if (SourceClass = "") OR (ObjectClass = "") then
		ShowSubmitTF = False
		ShowStr = "参数传递错误"
	else
		Set RsSoueceObj = Conn.Execute("Select * from FS_NewsClass where ClassID='" & SourceClass & "'")
		if RsSoueceObj.Eof then
			ShowSubmitTF = False
			ShowStr = "源栏目不存在"
			AllowOperation = False
		else
			ShowStr = "确定要把［" & RsSoueceObj("ClassCName") & "］合并到"
		end if
		Set RsObjectObj = Conn.Execute("Select * from FS_NewsClass where ClassID='" & ObjectClass & "'")
		if RsObjectObj.Eof then
			ShowSubmitTF = False
			ShowStr = "目标栏目不存在"
			AllowOperation = False
		else
			if AllowOperation = True then
				ShowStr = ShowStr & "［" & RsObjectObj("ClassCName") & "］吗？"
			end if
		end if
		if Result = "Submit" then
			if AllowOperation = True then
				Dim MergeSql
				MoveNewsFile "",SourceClass,ObjectClass
				MergeSql = "Update FS_News Set ClassID='" & ObjectClass & "' where ClassID in(" & AllClassID & ")"
				Conn.Execute(MergeSql)
				MergeSql = "Update FS_download Set ClassID='" & ObjectClass & "' where ClassID in(" & SourceClass & ")"
				Conn.Execute(MergeSql)
				MergeSql = "Update FS_Contribution set ClassID='" & ObjectClass & "' where ClassID in (" & AllClassID & ")"
				'------------/l
				DelClass(SourceClass)
				'---------------
				Set RsSoueceObj = Nothing
				Set RsObjectObj = Nothing
				Response.write("<script>window.close();</script>")
				Response.end
			else
				Set RsSoueceObj = Nothing
				Set RsObjectObj = Nothing
				Response.write("<script>alert('合并栏目不存在');window.close();</script>")
				Response.end
			end if
		end if
		Set RsSoueceObj = Nothing
		Set RsObjectObj = Nothing
	end if
Else
	ShowStr="源栏目与目标栏目相同，不可以合并！"
End If
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
      <td height="20" colspan="2"></td>
    </tr>
    <tr>
      <td width="28%" rowspan="3"><div align="center"><strong><font size="2"><img src="../../Images/Question.gif" width="39" height="37"></font></strong></div></td>
    </tr>
    <tr>
      <td width="72%"><% = ShowStr %></td>
    </tr>
    <tr>
      <td width="72%" height="20"><font color="#FF0000">合并后删除原来栏目(含其所有子栏目)</font>
<input name="DelSource" type="checkbox" value="Del"></td>
    </tr>
    <tr> 
<%
if ShowSubmitTF = true then
%>
      <td colspan="2" height="40"><div align="center"> 
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
Sub DelClass(DelClassID)

	Dim AllClassID,Sql,DelNewsSysRootDir,MyFile
	AllClassID = "'" & DelClassID & "'" & ChildClassIDList(DelClassID)
	If SysRootDir<>"" then 
		DelNewsSysRootDir="/"& SysRootDir
	else
		DelNewsSysRootDir=""
	End If
	Set MyFile=Server.CreateObject(G_FS_FSO)
	'---------------------物理文件删除-------------------------------------
	Dim DelClassFileObj
	Set DelClassFileObj = Conn.Execute("Select ClassEName,SaveFilePath from FS_NewsClass where ClassID in ("&AllClassID&")")
	Do while Not DelClassFileObj.eof
		If MyFile.FolderExists(Server.Mappath(DelNewsSysRootDir&DelClassFileObj("SaveFilePath")&"/"&DelClassFileObj("ClassEName"))) then
			MyFile.DeleteFolder(Server.Mappath(DelNewsSysRootDir&DelClassFileObj("SaveFilePath")&"/"&DelClassFileObj("ClassEName")))
		End if
		DelClassFileObj.MoveNext
	Loop
	DelClassFileObj.Close
	Set DelClassFileObj = Nothing
	set MyFile=Nothing
	'－－－－－－－－－－－－－－－－－－－－
	'Sql = "Delete from News where ClassID in (" & AllClassID & ")"
	'Conn.Execute(Sql)
	'if Err.Number <> 0 then Alert "删除栏目下的新闻失败"
	'Sql = "Delete from Contribution where ClassID in (" & AllClassID & ")"
	'Conn.Execute(Sql)
	'if Err.Number <> 0 then Alert "删除栏目下的投稿失败"
	'Sql = "Delete from DownLoad where ClassID in (" & AllClassID & ")"
	'Conn.Execute(Sql)
	'if Err.Number <> 0 then Alert "删除栏目下的下载失败"
	If request("DelSource")="Del" then 
		Sql = "Delete from FS_NewsClass where ClassID in (" & AllClassID & ")"
		Conn.Execute(Sql)
		if Err.Number <> 0 then Alert "删除栏目失败"
	End if
End Sub
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
%>
