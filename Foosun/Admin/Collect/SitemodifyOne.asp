<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="inc/Config.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<%
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
Dim DBC,Conn,CollectConn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
DBC.ConnStr = CollectDBConnectionStr
Set CollectConn = DBC.OpenConnection()
Set DBC = Nothing
'�ж�Ȩ��
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P080102") then Call ReturnError1()
'�ж�Ȩ�޽���
Dim SelectPath
if SysRootDir = "" then
	SelectPath = "/" & UpFiles
else
	SelectPath = "/" & SysRootDir & "/" & UpFiles
end if
Dim RsEditObj,EditSql,SiteID
Set RsEditObj = Server.CreateObject ("ADODB.RecordSet")
SiteID = Request("SiteID")
if SiteID <> "" then
	EditSql="Select * from FS_Site where ID=" & SiteID
	RsEditObj.Open EditSql,CollectConn,1,3
	if RsEditObj.Eof then
		Response.write"<script>alert(""û���޸ĵ�վ��"");location.href=""javascript:history.back()"";</script>"
		Response.end
	end if
else
	Response.write"<script>alert(""û���޸ĵ�վ��"");location.href=""javascript:history.back()"";</script>"
	Response.end
end if

Dim TempClassListStr
TempClassListStr = ClassList
Function ClassList()
	Dim ClassListObj,SelectStr
	Set ClassListObj = Conn.Execute("Select ClassID,ClassCName from FS_NewsClass where ParentID='0' order by ClassID desc")
	do while Not ClassListObj.Eof
		if RsEditObj("SysClass") = ClassListObj("ClassID") then
			SelectStr = "selected"
		else
			SelectStr = ""
		end if
		ClassList = ClassList & "<option " & SelectStr & " value="&ClassListObj("ClassID")&"" & ">" & ClassListObj("ClassCName") & "</option><br>"
		ClassList = ClassList & ChildClassList(ClassListObj("ClassID"),"")
		ClassListObj.MoveNext	
	loop
	ClassListObj.Close
	Set ClassListObj = Nothing
End Function

Function SiteFolderList()
	Dim ClassListObj,SelectStr
	Set ClassListObj = CollectConn.Execute("Select * from FS_SiteFolder order by ID desc")
	do while Not ClassListObj.Eof
		if RsEditObj("Folder") = ClassListObj("ID") then
			SelectStr = "selected"
		else
			SelectStr = ""
		end if
		SiteFolderList = SiteFolderList & "<option " & SelectStr & " value="&ClassListObj("ID")&"" & ">&nbsp;&nbsp;|--" & ClassListObj("SiteFolder") & "</option><br>"
		ClassListObj.MoveNext	
	loop
	ClassListObj.Close
	Set ClassListObj = Nothing
End Function

Function ChildClassList(ClassID,Temp)
	Dim TempRs,TempStr,SelectStr
	Set TempRs = Conn.Execute("Select ClassID,ClassCName from FS_NewsClass where ParentID='" & ClassID & "' order by ClassID desc")
	TempStr = Temp & " |- "
	do while Not TempRs.Eof
		if RsEditObj("SysClass") = TempRs("ClassID") then
			SelectStr = "selected"
		else
			SelectStr = ""
		end if
		ChildClassList = ChildClassList & "<option " & SelectStr & " value="&TempRs("ClassID")&"" & ">" & TempStr & TempRs("ClassCName") & "</option><br>"
		ChildClassList = ChildClassList & ChildClassList(TempRs("ClassID"),TempStr)
		TempRs.MoveNext
	loop
	TempRs.Close
	Set TempRs = Nothing
End Function

if Request.Form("Result") = "Edit" then
    Dim RsAddObj,sql
    if Request.Form("SaveIMGPath") = "" OR Request.Form("SiteName")="" Or Request.Form("SysTemplet")="" or Request.Form("objURL")="" or Request.Form("SysClass")=""  then
		Response.write"<script>alert(""����д������"");location.href=""javascript:history.back()"";</script>"
		Response.end
	end if
	Set RsAddObj = Server.CreateObject ("ADODB.RecordSet")
	Sql = "select * from FS_Site where id=" & Request.Form("SiteID")
	RsAddObj.Open Sql,CollectConn,1,3
	RsAddObj("SiteName") = NoCSSHackAdmin(Request.Form("SiteName"),"վ������")
	RsAddObj("objURL") = Request.Form("objURL")
	RsAddObj("SysClass") = Request.Form("SysClass")
	RsAddObj("SysTemplet") = Request.Form("SysTemplet")

On Error Resume Next 
Dim ListSetting,LinkSetting,PageBodySetting,PageTitleSetting,OtherNewsPageSetting,AuthorSetting,SourceSetting,AddDateSetting,OtherPageSetting,StrErr
	StrErr = ""
	ListSetting = Split(Request.Form("ListSetting"),"[�б�����]",-1,1)
	RsAddObj("ListHeadSetting") = ListSetting(0)
	RsAddObj("ListFootSetting") = ListSetting(1)
	If ListSetting(0)="" Or ListSetting(1)="" Or ListSetting(0)=Null Or ListSetting(1)=Null Or err Then
		If Err Then Err.clear
		RsAddObj("ListHeadSetting") = "<body"
		RsAddObj("ListFootSetting") = "</body>"
	End If
	LinkSetting = Split(Request.Form("LinkSetting"),"[�б�URL]",-1,1)
	RsAddObj("LinkHeadSetting") = LinkSetting(0)
	RsAddObj("LinkFootSetting") = LinkSetting(1)
	If err Then
		StrErr = "�б�URLû�����û����ò���ȷ��"
		Err.clear
	End if
	PageBodySetting = Split(Request.Form("PageBodySetting"),"[��������]",-1,1)
	RsAddObj("PagebodyHeadSetting") = PageBodySetting(0)
	RsAddObj("PagebodyFootSetting") = PageBodySetting(1)
	If err Then
		StrErr = StrErr & "\r\n��������û�����û����ò���ȷ��"
		Err.clear
	End if
	PageTitleSetting = Split(Request.Form("PageTitleSetting"),"[���ű���]",-1,1) 
	RsAddObj("PageTitleHeadSetting") = PageTitleSetting(0)
	RsAddObj("PageTitleFootSetting") = PageTitleSetting(1)
	If err Then
		StrErr = StrErr & "\r\n���ű���û�����û����ò���ȷ��"
		Err.clear
	End If
	If InStr(Request.Form("OtherNewsPageSetting"),"[��ҳ����]")<>0 Then
		OtherNewsPageSetting = Split(Request.Form("OtherNewsPageSetting"),"[��ҳ����]",-1,1)
		RsAddObj("OtherNewsPageHeadSetting") = OtherNewsPageSetting(0)
		RsAddObj("OtherNewsPageFootSetting") = OtherNewsPageSetting(1)
	End if
	If InStr(Request.Form("AuthorSetting"),"[����]")<>0 then
		AuthorSetting = Split(Request.Form("AuthorSetting"),"[����]",-1,1)
		RsAddObj("AuthorHeadSetting") = AuthorSetting(0)
		RsAddObj("AuthorFootSetting") = AuthorSetting(1)
	End If
	If InStr(Request.Form("SourceSetting"),"[��Դ]")<>0 then
		SourceSetting = Split(Request.Form("SourceSetting"),"[��Դ]",-1,1)
		RsAddObj("SourceHeadSetting") = SourceSetting(0)
		RsAddObj("SourceFootSetting") = SourceSetting(1)
	End If
	If InStr(Request.Form("AddDateSetting"),"[����ʱ��]")<>0 then
		AddDateSetting = Split(Request.Form("AddDateSetting"),"[����ʱ��]",-1,1)
		RsAddObj("AddDateHeadSetting") = AddDateSetting(0)
		RsAddObj("AddDateFootSetting") = AddDateSetting(1)
	End if
	If StrErr<>"" Then
		Err.clear
		Response.Write "<script>alert('"& StrErr &"');history.back();</script>"
		Response.End
	End If 
	RsAddObj("SaveIMGPath") = Request.Form("SaveIMGPath")
	Select Case Request.Form("OtherType")
		Case "0"
			RsAddObj("OtherType") = Request.Form("OtherType")
			RsAddObj("IndexRule") = ""
			RsAddObj("StartPageNum") = ""
			RsAddObj("EndPageNum") = ""
			RsAddObj("HandPageContent") = ""
			RsAddObj("OtherPageHeadSetting") = ""
			RsAddObj("OtherPageFootSetting") = ""
		Case "1"
			RsAddObj("OtherType") = Request.Form("OtherType")
			RsAddObj("IndexRule") = ""
			RsAddObj("StartPageNum") = ""
			RsAddObj("EndPageNum") = ""
			RsAddObj("HandPageContent") = ""
			OtherPageSetting = Split(Request.Form("OtherPageSetting"),"[����ҳ��]",-1,1)
			RsAddObj("OtherPageHeadSetting") = OtherPageSetting(0)
			RsAddObj("OtherPageFootSetting") = OtherPageSetting(1)
		Case "2"
			RsAddObj("OtherType") = Request.Form("OtherType")
			RsAddObj("IndexRule") = Request.Form("IndexRule")
			RsAddObj("StartPageNum") = Request.Form("StartPageNum")
			RsAddObj("EndPageNum") = Request.Form("EndPageNum")
			RsAddObj("HandPageContent") = ""
			RsAddObj("OtherPageHeadSetting") = ""
			RsAddObj("OtherPageFootSetting") = ""
		Case "3"
			RsAddObj("OtherType") = Request.Form("OtherType")
			RsAddObj("IndexRule") = ""
			RsAddObj("StartPageNum") = ""
			RsAddObj("EndPageNum") = ""
			RsAddObj("HandPageContent") = Request.Form("HandPageContent")
			RsAddObj("OtherPageHeadSetting") = ""
			RsAddObj("OtherPageFootSetting") = ""
		Case Else
			RsAddObj("OtherType") = 0
			RsAddObj("IndexRule") = ""
			RsAddObj("StartPageNum") = ""
			RsAddObj("EndPageNum") = ""
			RsAddObj("HandPageContent") = ""
			RsAddObj("OtherPageHeadSetting") = ""
			RsAddObj("OtherPageFootSetting") = ""
	End Select
	RsAddObj("HandSetAuthor") = Request.Form("HandSetAuthor")
	RsAddObj("HandSetSource") = Request.Form("HandSetSource")
	if IsDate(Request.Form("HandSetAddDate")) then
		RsAddObj("HandSetAddDate") = Request.Form("HandSetAddDate")
	end if
	if Request.Form("IsIFrame") = "1" then
		RsAddObj("IsIFrame") = True
	else
		RsAddObj("IsIFrame") = False
	end if
	if Request.Form("IsScript") = "1" then
		RsAddObj("IsScript") = True
	else
		RsAddObj("IsScript") = False
	end if
	if Request.Form("IsClass") = "1" then
		RsAddObj("IsClass") = True
	else
		RsAddObj("IsClass") = False
	end if
	if Request.Form("IsFont") = "1" then
		RsAddObj("IsFont") = True
	else
		RsAddObj("IsFont") = False
	end if
	if Request.Form("IsSpan") = "1" then
		RsAddObj("IsSpan") = True
	else
		RsAddObj("IsSpan") = False
	end if
	if Request.Form("IsObject") = "1" then
		RsAddObj("IsObject") = True
	else
		RsAddObj("IsObject") = False
	end if
	if Request.Form("IsStyle") = "1" then
		RsAddObj("IsStyle") = True
	else
		RsAddObj("IsStyle") = False
	end if
	if Request.Form("IsDiv") = "1" then
		RsAddObj("IsDiv") = True
	else
		RsAddObj("IsDiv") = False
	end if
	if Request.Form("IsA") = "1" then
		RsAddObj("IsA") = True
	else
		RsAddObj("IsA") = False
	end if
	if Request.Form("Audit") = "1" then
		RsAddObj("Audit") = True
	else
		RsAddObj("Audit") = False
	end if
	if Request.Form("IsAutoCollect") <> "" then
		RsAddObj("IsAutoCollect") = True
		RsAddObj("CollectDate") = Clng(Request.Form("CollectDate"))
	else
		RsAddObj("IsAutoCollect") = False
	end if
	if Request.Form("TextTF") = "1" then
		RsAddObj("TextTF") = True
	else
		RsAddObj("TextTF") = False
	end If
	If Request.Form("IsReverse") = "1" Then
		RsAddObj("IsReverse") = "1"
	Else
		RsAddObj("IsReverse") = "0"
	End If
	if Request.Form("SaveRemotePic") = "1" then
		RsAddObj("SaveRemotePic") = True
	else
		RsAddObj("SaveRemotePic") = False
	end if
	if Request.Form("Islock") <> "" then
		RsAddObj("Islock") = True
	else
		RsAddObj("Islock") = False
	end if
	RsAddObj.update
	RsAddObj.close
	Set RsAddObj = Nothing
	Response.Redirect("Site.asp")
	Response.End
end if

Dim TempletDirectory
if SysRootDir <> "" then
	TempletDirectory = "/" & SysRootDir & "/" & TempletDir
else
	TempletDirectory = "/" & TempletDir
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�Զ����Ųɼ���վ������</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<body leftmargin="2" topmargin="2">
<form name="Form" method="post" action="" id="Form">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td width="30" align="center" alt="����" onClick="document.Form.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
			<td width=2 class="Gray">|</td>
		    <td width="35" align="center" alt="����" onClick="history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
            <td>&nbsp; <input name="vs" type="hidden" id="vs2" value="add"> <input name="SiteID" type="hidden" id="SiteID2" value="<% = SiteID %>"> 
              <input name="Result" type="hidden" id="Result2" value="Edit"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1">
    <tr> 
      <td width="15%" height="26"> <div align="center">�ɼ�վ������</div></td>
      <td> <input name="SiteName" style="width:100%;" type="text" id="SiteName" value="<%=RsEditObj("sitename")%>"> 
        <div align="right"> </div></td>
    </tr>
    <tr> 
      <td height="26"> <div align="center">�ɼ�����ҳ</div></td>
      <td><input name="objURL" style="width:100%;" type="text" id="objURL" value="<%=RsEditObj("objURL")%>" size="50"></td>
    </tr>
	<tr> 
		<td height="26"> <div align="center">�ɼ�վ�����</div></td>
      <td><select name="SiteFolder" style="width:100%;" id="SiteFolder">
		<option value="0">����Ŀ</option>
          <% = SiteFolderList %>
        </select></td>
    </tr>
    <tr> 
      <td height="26"><div align="center">���Ŀ����Ŀ</div></td>
      <td><select name="SysClass" style="width:100%;" id="SysClass">
          <% = TempClassListStr %>
        </select></td>
    </tr>
    <tr> 
      <td height="26"> <div align="center">��������</div></td>
      <td><input readonly name="SysTemplet" type="text" id="SysTemplet" style="width:80%;" value="<%=RsEditObj("SysTemplet")%>"> 
        <input name="Submitaaa" type="button" id="Submitaaa" value="ѡ��ģ��" onClick="OpenWindowAndSetValue('../../FunPages/SelectFileFrame.asp?CurrPath=<% = TempletDirectory %>',400,300,window,document.Form.SysTemplet);"> 
        <div align="right"></div></td>
    </tr>
    <tr> 
      <td height="26"><div align="center">�ɼ�����</div></td>
      <td>���� 
        <input name="islock" type="checkbox" id="islock" value="1" <%if RsEditObj("islock")=true then response.Write("checked")%>>
        ����Զ��ͼƬ 
        <input type="checkbox" name="SaveRemotePic" value="1" <%if RsEditObj("SaveRemotePic")=true then response.Write("checked")%>>
        �����Ƿ��Ѿ���� 
        <input type="checkbox" name="Audit" value="1" <%if RsEditObj("Audit")=true then response.Write("checked")%>>
		�Ƿ���ɼ� 
        <input name="IsReverse" type="checkbox" id="IsReverse" value="1" <%if RsEditObj("IsReverse")="1" then response.Write("checked")%>>
	</td>
    </tr>
    <tr>
      <td height="26"><div align="center">����ͼƬ·��</div></td>
      <td><input type="text" readonly name="SaveIMGPath" style="width:80%;" value="<% = RsEditObj("SaveIMGPath") %>">
        <input name="Submit111" id="SelectPath" type="button" value="ѡ��·��" onClick="OpenWindowAndSetValue('../../FunPages/SelectPathFrame.asp?CurrPath=<% = SelectPath %>',400,300,window,document.Form.SaveIMGPath);"></td>
    </tr>
    <tr> 
      <td height="26"><div align="center">����ѡ��</div></td>
      <td>HTML <input type="checkbox" name="TextTF" value="1" <% if RsEditObj("TextTF") = True then Response.Write("checked")%>>
        STYLE <input type="checkbox" name="IsStyle" value="1" <% if RsEditObj("IsStyle") = True then Response.Write("checked")%>>
        DIV<input type="checkbox" name="IsDiv" value="1" <% if RsEditObj("IsDiv") = True then Response.Write("checked")%>>
        A<input type="checkbox" name="IsA" value="1" <% if RsEditObj("IsA") = True then Response.Write("checked")%>>
        CLASS<input type="checkbox" name="IsClass" value="1" <% if RsEditObj("IsClass") = True then Response.Write("checked")%>>
        FONT<input type="checkbox" name="IsFont" value="1" <% if RsEditObj("IsFont") = True then Response.Write("checked")%>>
        SPAN<input type="checkbox" name="IsSpan" value="1" <% if RsEditObj("IsSpan") = True then Response.Write("checked")%>>
        OBJECT<input type="checkbox" name="IsObject" value="1" <% if RsEditObj("IsObject") = True then Response.Write("checked")%>>
        IFRAME<input type="checkbox" name="IsIFrame" value="1" <% if RsEditObj("IsIFrame") = True then Response.Write("checked")%>>
        SCRIPT<input type="checkbox" name="IsScript" value="1" <% if RsEditObj("IsScript") = True then Response.Write("checked")%>>
        </td>
    </tr>
    <tr> 
      <td height="36" colspan="2">
<div align="center"></div>
        <div align="center">
          <input onClick="ChangeCutPara(0);" <% if RsEditObj("OtherType") = 0 then Response.Write("checked") %> name="OtherType" type="radio" value="0">
          ����ҳ 
          <input type="radio" onClick="ChangeCutPara(1);" name="OtherType" <% if RsEditObj("OtherType") = 1 then Response.Write("checked") %> value="1">
          ��Ƿ�ҳ���� 
          <input type="radio" onClick="ChangeCutPara(2);" <% if RsEditObj("OtherType") = 2 then Response.Write("checked") %> name="OtherType" value="2">
          ������ҳ���� 
          <input type="radio" onClick="ChangeCutPara(3);" <% if RsEditObj("OtherType") = 3 then Response.Write("checked") %> name="OtherType" value="3">
          �ֹ���ҳ����
		  <input type="radio" onClick="ChangeCutPara(4);" <% if RsEditObj("OtherType") = 4 then Response.Write("checked") %> name="OtherType" value="4">
          <b>�б����ݷ�Χ����</b></div></td>
    </tr>
    <tr id="TagCutPage" style="display:<% if RsEditObj("OtherType") <> 1 then Response.Write("none") %>;"> 
      <td colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="15%"> 
              <div align="center">����ҳ��</div></td>
            <td>
			&nbsp;&nbsp;��������
			<span onClick="if(document.Form.OtherPageSetting.rows>2)document.Form.OtherPageSetting.rows-=1" style='cursor:hand'><b>��С</b></span>
			<span onClick="document.Form.OtherPageSetting.rows+=1" style='cursor:hand'><b>����</b></span>
			&nbsp;&nbsp;���ñ�ǩ��<font onmouseover="getActiveText(document.Form.OtherPageSetting);" onClick="addTag('[����ҳ��]')" style="CURSOR: hand"><b>[����ҳ��]</b></font>
			&nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.Form.OtherPageSetting);" onClick="addTag('[����]')" style="CURSOR: hand"><b>[����]</b></font>
			<br>
			<textarea  ondblclick="getActiveText(this)" onclick="getActiveText(this)"  onchange="getActiveText(this)" name="OtherPageSetting" id="OtherPageSetting" rows="4" style="width:100%;"><%=RsEditObj("OtherPageHeadSetting")%>[����ҳ��]<%=RsEditObj("OtherPageFootSetting")%></textarea></td>
          </tr>
        </table></td>
    </tr>
    <tr id="IndexCutPage" style="display:<% if RsEditObj("OtherType") <> 2 then Response.Write("none") %>;"> 
      <td colspan="2"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="15%"> 
              <div align="center">�������� </div></td>
            <td>
			&nbsp;&nbsp;��������
			<span onClick="if(document.Form.IndexRule.rows>2)document.Form.IndexRule.rows-=1" style='cursor:hand'><b>��С</b></span>
			<span onClick="document.Form.IndexRule.rows+=1" style='cursor:hand'><b>����</b></span><br>
			<textarea name="IndexRule" rows="3" id="IndexRule" style="width:100%;"><% = RsEditObj("IndexRule") %></textarea></td>
          </tr>
          <tr> 
            <td height="26"> <div align="center">ҳ��</div></td>
            <td>ҳ�뿪ʼ�� 
              <input name="StartPageNum" type="text" id="StartPageNum" size="3" maxlength="8" value="<% = RsEditObj("StartPageNum") %>">
              ҳ����� 
              <input name="EndPageNum" type="text" id="EndPageNum" size="3" maxlength="8" value="<% = RsEditObj("EndPageNum") %>">&nbsp&nbsp��:������������дhttp://.../index_^$^.htm������^$^�����趨��ҳ��</td>
          </tr>
        </table></td>
    </tr>
    <tr id="HandCutPage" style="display:<% if RsEditObj("OtherType") <> 3 then Response.Write("none") %>;"> 
      <td width="10%"> <div align="center">��ҳ����</div></td>
      <td height="26">	  &nbsp;&nbsp;��������
			<span onClick="if(document.Form.HandPageContent.rows>2)document.Form.HandPageContent.rows-=1" style='cursor:hand'><b>��С</b></span>
			<span onClick="document.Form.HandPageContent.rows+=1" style='cursor:hand'><b>����</b></span>
			<textarea  name="HandPageContent" rows="6" id="HandPageContent" style="width:100%;"><% = RsEditObj("HandPageContent") %></textarea></tr>
    <tr  id="ListContent" style="display:none"> 
      <td colspan="2">
	  <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="15%"> 
	  <div align="center">�б�����</div></td>
      <td>	&nbsp;&nbsp;��������
			<span onClick="if(document.Form.ListSetting.rows>2)document.Form.ListSetting.rows-=1" style='cursor:hand'><b>��С</b></span>
			<span onClick="document.Form.ListSetting.rows+=1" style='cursor:hand'><b>����</b></span>
	  &nbsp;&nbsp;���ñ�ǩ:<font onmouseover="getActiveText(document.Form.ListSetting);" onClick="addTag('[�б�����]')" style="CURSOR: hand"><b>[�б�����]</b></font>
	  &nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.Form.ListSetting);" onClick="addTag('[����]')" style="CURSOR: hand"><b>[����]</b></font><br>
	   <textarea  ondblclick="getActiveText(this)" onclick="getActiveText(this)"  onchange="getActiveText(this)" name="ListSetting" cols="50" rows="3" id="ListSetting" style="width:100%;"><%=RsEditObj("ListHeadSetting")%>[�б�����]<%=RsEditObj("ListFootSetting")%></textarea>
	   </td>
          </tr>
        </table>
	   </td>
    </tr>
    <tr> 
      <td> <div align="center">�б�URL<font color="#ff0000">*</font></div></td>
      <td>	&nbsp;&nbsp;��������
			<span onClick="if(document.Form.LinkSetting.rows>2)document.Form.LinkSetting.rows-=1" style='cursor:hand'><b>��С</b></span>
			<span onClick="document.Form.LinkSetting.rows+=1" style='cursor:hand'><b>����</b></span>
			&nbsp;&nbsp;���ñ�ǩ��<font onmouseover="getActiveText(document.Form.LinkSetting);" onClick="addTag('[�б�URL]')" style="CURSOR: hand"><b>[�б�URL]</b></font>&nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.Form.LinkSetting);" onClick="addTag('[����]')" style="CURSOR: hand"><b>[����]</b></font><br>
	  <textarea   ondblclick="getActiveText(this)" onclick="getActiveText(this)"  onchange="getActiveText(this)"  name="LinkSetting" cols="50" rows="3" id="textarea2" style="width:100%;"><%=RsEditObj("LinkHeadSetting")%>[�б�URL]<%=RsEditObj("LinkFootSetting")%></textarea></td>
    </tr>
    <tr> 
      <td> <div align="center">���ű���<font color="#ff0000">*</font></div></td>
      <td>	&nbsp;&nbsp;��������
			<span onClick="if(document.Form.PageTitleSetting.rows>2)document.Form.PageTitleSetting.rows-=1" style='cursor:hand'><b>��С</b></span>
			<span onClick="document.Form.PageTitleSetting.rows+=1" style='cursor:hand'><b>����</b></span>
			&nbsp;&nbsp;���ñ�ǩ��<font onmouseover="getActiveText(document.Form.PageTitleSetting);" onClick="addTag('[���ű���]')" style="CURSOR: hand"><b>[���ű���]</b></font>&nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.Form.PageTitleSetting);" onClick="addTag('[����]')" style="CURSOR: hand"><b>[����]</b></font><br>
	  <textarea  ondblclick="getActiveText(this)" onclick="getActiveText(this)"  onchange="getActiveText(this)"  name="PageTitleSetting" cols="50" rows="3" id="textarea6" style="width:100%;"><%=RsEditObj("PageTitleHeadSetting")%>[���ű���]<%=RsEditObj("PageTitleFootSetting")%></textarea></td>
    </tr>
    <tr> 
      <td> <div align="center">��������<font color="#ff0000">*</font></div></td>
      <td>	&nbsp;&nbsp;��������
			<span onClick="if(document.Form.PagebodySetting.rows>2)document.Form.PagebodySetting.rows-=1" style='cursor:hand'><b>��С</b></span>
			<span onClick="document.Form.PagebodySetting.rows+=1" style='cursor:hand'><b>����</b></span>
	  &nbsp;&nbsp;���ñ�ǩ��<font onmouseover="getActiveText(document.Form.PagebodySetting);" onClick="addTag('[��������]')" style="CURSOR: hand"><b>[��������]</b></font>&nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.Form.PagebodySetting);" onClick="addTag('[����]')" style="CURSOR: hand"><b>[����]</b></font><br>
	   <textarea  ondblclick="getActiveText(this)" onclick="getActiveText(this)"  onchange="getActiveText(this)"  name="PagebodySetting" cols="50" rows="3" id="textarea8" style="width:100%;"><%=RsEditObj("PagebodyHeadSetting")%>[��������]<%=RsEditObj("PagebodyFootSetting")%></textarea></td>
    </tr>
    <tr> 
      <td height="26" colspan="4"> <div align="center">
          <input name="OtherSetType" type="radio" onClick="ChangeSetOption(0);" value="0" checked>
          �������� 
          <input type="radio" name="OtherSetType" onClick="ChangeSetOption(1);" value="1">
          ������Դ 
          <input type="radio" name="OtherSetType" onClick="ChangeSetOption(2);" value="2">
          ����ʱ�� 
          <input type="radio" name="OtherSetType" onClick="ChangeSetOption(3);" value="3">
          ���÷�ҳ </div></td>
    </tr>
    <tr id="SetAuthor" style="display:;"> 
      <td colspan="4"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="26"> 
              <div align="center">�ֶ�����</div></td>
            <td colspan="3"><input style="width:100%;" type="text" name="HandSetAuthor" value="<% = RsEditObj("HandSetAuthor") %>"></td>
          </tr>
          <tr> 
            <td width="15%"> 
              <div align="center">����</div></td>
            <td colspan="3">	&nbsp;&nbsp;��������
			<span onClick="if(document.Form.AuthorSetting.rows>2)document.Form.AuthorSetting.rows-=1" style='cursor:hand'><b>��С</b></span>
			<span onClick="document.Form.AuthorSetting.rows+=1" style='cursor:hand'><b>����</b></span>
			&nbsp;&nbsp;���ñ�ǩ��<font onmouseover="getActiveText(document.Form.AuthorSetting);" onClick="addTag('[����]')" style="CURSOR: hand"><b>[����]</b></font>&nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.Form.AuthorSetting);" onClick="addTag('[����]')" style="CURSOR: hand"><b>[����]</b></font><br>
			<textarea  ondblclick="getActiveText(this)" onclick="getActiveText(this)"  onchange="getActiveText(this)"  name="AuthorSetting" cols="50" rows="3" id="textarea9" style="width:100%;"><%=RsEditObj("AuthorHeadSetting")%>[����]<%=RsEditObj("AuthorFootSetting")%></textarea></td>
          </tr>
        </table></td>
    </tr>
    <tr id="SetSource" style="display:none;"> 
      <td colspan="4"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="26">
<div align="center">�ֶ�����</div></td>
            <td colspan="3"><input style="width:100%;" type="text" name="HandSetSource" value="<% = RsEditObj("HandSetSource") %>"></td>
          </tr>
		  <tr> 
            <td width="15%"> 
              <div align="center">��Դ</div></td>
            <td colspan="3">	&nbsp;&nbsp;��������
			<span onClick="if(document.Form.SourceSetting.rows>2)document.Form.SourceSetting.rows-=1" style='cursor:hand'><b>��С</b></span>
			<span onClick="document.Form.SourceSetting.rows+=1" style='cursor:hand'><b>����</b></span>
			&nbsp;&nbsp;���ñ�ǩ��<font onmouseover="getActiveText(document.Form.SourceSetting);" onClick="addTag('[��Դ]')" style="CURSOR: hand"><b>[��Դ]</b></font>&nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.Form.SourceSetting);" onClick="addTag('[����]')" style="CURSOR: hand"><b>[����]</b></font><br>
			 <textarea   ondblclick="getActiveText(this)" onclick="getActiveText(this)"  onchange="getActiveText(this)" name="SourceSetting" cols="50" rows="3" id="textarea9a" style="width:100%;"><%=RsEditObj("SourceHeadSetting")%>[��Դ]<%=RsEditObj("SourceFootSetting")%></textarea></td>
          </tr>
        </table></td>
    </tr>
    <tr id="SetAddTime" style="display:none;"> 
      <td colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="26">
<div align="center">�ֶ�����</div></td>
            <td colspan="3"><input style="width:100%;" type="text" name="HandSetAddDate" value="<% = RsEditObj("HandSetAddDate") %>"></td>
          </tr>
		  <tr> 
            <td width="15%"> 
              <div align="center">����ʱ��</div></td>
            <td>	&nbsp;&nbsp;��������
			<span onClick="if(document.Form.AddDateSetting.rows>2)document.Form.AddDateSetting.rows-=1" style='cursor:hand'><b>��С</b></span>
			<span onClick="document.Form.AddDateSetting.rows+=1" style='cursor:hand'><b>����</b></span>
			&nbsp;&nbsp;���ñ�ǩ��<font onmouseover="getActiveText(document.Form.AddDateSetting);" onClick="addTag('[����ʱ��]')" style="CURSOR: hand"><b>[����ʱ��]</b></font>&nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.Form.AddDateSetting);" onClick="addTag('[����]')" style="CURSOR: hand"><b>[����]</b></font><br>
			 <textarea  ondblclick="getActiveText(this)" onclick="getActiveText(this)"  onchange="getActiveText(this)"  name="AddDateSetting" cols="50" rows="3" id="textarea9" style="width:100%;"><%=RsEditObj("AddDateHeadSetting")%>[����ʱ��]<%=RsEditObj("AddDateFootSetting")%></textarea></td>
          </tr>
        </table></td>
    </tr>
    <tr id="SetCutPage" style="display:none;"> 
      <td colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr> 
            <td width="15%"> 
              <div align="center">��ҳ����<br>(��һҳ)</div></td>
      <td> 	&nbsp;&nbsp;��������
			<span onClick="if(document.Form.OtherNewsPageSetting.rows>2)document.Form.OtherNewsPageSetting.rows-=1" style='cursor:hand'><b>��С</b></span>
			<span onClick="document.Form.OtherNewsPageSetting.rows+=1" style='cursor:hand'><b>����</b></span>
	  &nbsp;&nbsp;���ñ�ǩ��<font onmouseover="getActiveText(document.Form.OtherNewsPageSetting);" onClick="addTag('[��ҳ����]')" style="CURSOR: hand"><b>[��ҳ����]</b></font>&nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.Form.OtherNewsPageSetting);" onClick="addTag('[����]')" style="CURSOR: hand"><b>[����]</b></font><br>
	  <textarea  ondblclick="getActiveText(this)" onclick="getActiveText(this)"  onchange="getActiveText(this)" name="OtherNewsPageSetting" cols="50" rows="3" id="textarea5" style="width:100%;"><%=RsEditObj("OtherNewsPageHeadSetting")%>[��ҳ����]<%=RsEditObj("OtherNewsPageFootSetting")%></textarea></td>
    </tr>
        </table></td>
    </tr>
</table>
</form>
<p><br><p><p>
</body>
</html>
<%
Set Conn = Nothing
Set CollectConn = Nothing
Set RsEditObj = Nothing
%>
<script language="JavaScript">
function ChangeCutPara(Flag)
{
	switch (Flag)
	{
		case 0 :
			document.all.TagCutPage.style.display='none';
			document.all.IndexCutPage.style.display='none';
			document.all.HandCutPage.style.display='none';
			document.all.ListContent.style.display='none';
			break;
		case 1 :
			document.all.TagCutPage.style.display='';
			document.all.IndexCutPage.style.display='none';
			document.all.HandCutPage.style.display='none';
			document.all.ListContent.style.display='none';
			break;
		case 2 :
			document.all.TagCutPage.style.display='none';
			document.all.IndexCutPage.style.display='';
			document.all.HandCutPage.style.display='none';
			document.all.ListContent.style.display='none';
			break;
		case 3 :
			document.all.TagCutPage.style.display='none';
			document.all.IndexCutPage.style.display='none';
			document.all.HandCutPage.style.display='';
			document.all.ListContent.style.display='none';
			break;
		case 4 :
			document.all.TagCutPage.style.display='none';
			document.all.IndexCutPage.style.display='none';
			document.all.HandCutPage.style.display='none';
			document.all.ListContent.style.display='';
			break;		
		default :
			document.all.TagCutPage.style.display='none';
			document.all.IndexCutPage.style.display='none';
			document.all.HandCutPage.style.display='none';
			document.all.ListContent.style.display='none';
			break;
	}
}
function ChangeSetOption(Flag)
{
	switch (Flag)
	{
		case 0 :
			document.all.SetAuthor.style.display='';
			document.all.SetSource.style.display='none';
			document.all.SetAddTime.style.display='none';
			document.all.SetCutPage.style.display='none';
			break;
		case 1 :
			document.all.SetAuthor.style.display='none';
			document.all.SetSource.style.display='';
			document.all.SetAddTime.style.display='none';
			document.all.SetCutPage.style.display='none';
			break;
		case 2 :
			document.all.SetAuthor.style.display='none';
			document.all.SetSource.style.display='none';
			document.all.SetAddTime.style.display='';
			document.all.SetCutPage.style.display='none';
			break;
		case 3 :
			document.all.SetAuthor.style.display='none';
			document.all.SetSource.style.display='none';
			document.all.SetAddTime.style.display='none';
			document.all.SetCutPage.style.display='';
			break;
		case 999 :
			document.all.SetAuthor.style.display='none';
			document.all.SetSource.style.display='none';
			document.all.SetAddTime.style.display='none';
			document.all.SetCutPage.style.display='none';
			break;
		default :
			document.all.SetAuthor.style.display='none';
			document.all.SetSource.style.display='none';
			document.all.SetAddTime.style.display='none';
			document.all.SetCutPage.style.display='none';
			break;
	}
}

currObj = "uuuu";
function getActiveText(obj)
{
	obj.focus();
	currObj = obj;
}

function addTag(code)
{
	addText(code);
}

function addText(ibTag)
{
	var isClose = false;
	var obj_ta = currObj;
//alert("ok");
	if (obj_ta.isTextEdit)
	{
	//alert("nooooo");
		obj_ta.focus();
		var sel = document.selection;
		var rng = sel.createRange();
		rng.colapse;

		if((sel.type == "Text" || sel.type == "None") && rng != null)
		{
			rng.text = ibTag;
		}

		obj_ta.focus();

		return isClose;
	}
	else
		return false;
}	
-->
</script>