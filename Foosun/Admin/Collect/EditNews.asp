<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="inc/Config.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
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
if Not JudgePopedomTF(Session("Name"),"P080301") then Call ReturnError1()
'�ж�Ȩ�޽���
Dim NewsIDStr,Result,RsNewsObj,Sql
Dim Title,Links,Content,AddDate,ClassID,SiteID,SysTemplet,Author,SourceStr
Result = Request("Result")
NewsIDStr = Request("NewsIDStr")
if Result = "Submit" then
	Title = Request.Form("Title")
	Links = Request.Form("Links")
	Content = Request.Form("Content")
	AddDate = Request.Form("AddDate")
	ClassID = Request.Form("ClassID")
	SiteID = Request.Form("SiteID")
	SysTemplet = Request.Form("SysTemplet")
	Author = Request.Form("Author")
	SourceStr = Request.Form("Source")
	if NewsIDStr <> "" then
		Sql = "Select * from FS_News where ID=" & NewsIDStr
		'On Error Resume Next
		Set RsNewsObj = Server.CreateObject("ADODB.RecordSet")
		RsNewsObj.Open Sql,CollectConn,3,3
		RsNewsObj("Title") = Title 
		RsNewsObj("Links") = Links
		RsNewsObj("Content") = Content
		RsNewsObj("AddDate") = AddDate
		RsNewsObj("ClassID") = ClassID
		RsNewsObj("SysTemplet") = SysTemplet
		RsNewsObj("Author") = Author
		RsNewsObj("Source") = SourceStr
		RsNewsObj("SiteID") = SiteID
		RsNewsObj.UpDate
		RsNewsObj.Close
		Set RsNewsObj = Nothing
		if Err.Number <> 0 then
%>
	<script language="JavaScript">
	alert('�޸�ʧ��');
	</script>
<%
		else
			Response.Redirect("Check.asp")
		end if
	else
%>
	<script language="JavaScript">
	alert('�޸ĵ����Ų�����');
	</script>
<%
	end if
else
	if NewsIDStr <> "" then
		Sql = "Select * from FS_News where ID=" & NewsIDStr
		Set RsNewsObj = CollectConn.Execute(Sql)
		if Not RsNewsObj.Eof then
			Title = RsNewsObj("Title")
			Links = RsNewsObj("Links")
			Content = RsNewsObj("Content")
			AddDate = RsNewsObj("AddDate")
			ClassID = RsNewsObj("ClassID")
			SiteID = RsNewsObj("SiteID")
			SysTemplet = RsNewsObj("SysTemplet")
			Author = RsNewsObj("Author")
			SourceStr = RsNewsObj("Source")
		else
%>
	<script language="JavaScript">
	alert('���Ų�����');
	</script>
<%
		end if
	else
%>
	<script language="JavaScript">
	alert('��������');
	</script>
<%
	end if
end if

Dim TempClassListStr
TempClassListStr = ClassList
Function ClassList()
	Dim ClassListObj,SelectStr
	Set ClassListObj = Conn.Execute("Select ClassID,ClassCName from FS_NewsClass where ParentID='0' order by ClassID desc")
	do while Not ClassListObj.Eof
		if ClassListObj("ClassID") = ClassID then
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

Function ChildClassList(ClassID,Temp)
	Dim TempRs,TempStr,SelectStr
	Set TempRs = Conn.Execute("Select ClassID,ClassCName from FS_NewsClass where ParentID='" & ClassID & "' order by ClassID desc")
	TempStr = Temp & " |- "
	do while Not TempRs.Eof
		if TempRs("ClassID") = ClassID then
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

Dim SiteList,RsSiteObj
Set RsSiteObj = Server.CreateObject("Adodb.RecordSet")
RsSiteObj.Source = "Select ID,SiteName from FS_Site order by id desc"
RsSiteObj.open RsSiteObj.Source,CollectConn,1,3
do while Not RsSiteObj.Eof
	if Clng(RsSiteObj("ID")) = Clng(SiteID) then
		SiteList = SiteList & "<option selected value=" & RsSiteObj("ID") & "" & ">" & RsSiteObj("SiteName") & "</option><br>"
	else
		SiteList = SiteList & "<option value=" & RsSiteObj("ID") & "" & ">" & RsSiteObj("SiteName") & "</option><br>"
	end if
	RsSiteObj.MoveNext	
loop
RsSiteObj.Close
Set RsSiteObj = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�޸�����</title>
</head>
<style>
.LableWindow
{
	border-right: 1px solid;
	border-left: 1px solid;
	border-bottom: 1px solid;
	border-color: Black;
	cursor: default;
}
.LableDefault
{
	border-right: 1px solid;
	border-top: 1px solid;
	font-size: 12px;
	border-left: 1px solid;
	border-bottom: 1px solid;
	font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
	border-color: Black;
	cursor: default;

}
.LableSelected
{
	border-right: 1px solid;
	border-top: 1px solid;
	font-size: 12px;
	border-left: 1px solid;
	font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
	border-color: Black;
	cursor: default;

}
.ToolBarButtonLine {
	border-bottom: 1px solid;
	font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
	border-color: Black;
}
</style>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<body topmargin="2" leftmargin="2">
<form action="" method="post" name="NewsForm">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td width="35" align="center" alt="����" onClick="document.NewsForm.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
			<td width=2 class="Gray">|</td>
			<td width="35" align="center" alt="����" onClick="history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
            <td>&nbsp; <input name="Result" type="hidden" id="Result" value="Submit"> 
              <input value="<% = NewsIDStr %>" name="NewsIDStr" type="hidden" id="NewsIDStr"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="120" height="26"> 
              <div align="center">���ű���</div></td>
            <td> <input name="Title" type="text" id="Title2" style="width:100%;" value="<% = Title %>"></td>
          </tr>
          <tr> 
            <td height="26"> 
              <div align="center">��������</div></td>
            <td> <div align="right"></div>
              <input name="Links" type="text" id="Links" style="width:100%;" value="<% = Links %>"></td>
          </tr>
          <tr> 
            <td height="26"> 
              <div align="center">Ŀ����Ŀ</div></td>
            <td> <select style="width:100%;" name="ClassID">
                <% = TempClassListStr %>
              </select></td>
          </tr>
          <tr> 
            <td height="26"> 
              <div align="center">��������</div></td>
            <td> <input name="SysTemplet" type="text" id="SysTemplet2" style="width:100%;" value="<% = SysTemplet %>"> 
              <div align="center"></div></td>
          </tr>
          <tr> 
            <td><div align="center">�ɼ�վ��</div></td>
            <td><select style="width:100%;" name="SiteID">
                <% = SiteList %>
              </select></td>
          </tr>
          <tr> 
            <td height="26">
<div align="center">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</div></td>
            <td><input style="width:100%;" type="text" name="Author" value="<% = Author %>"></td>
          </tr>
          <tr> 
            <td height="26">
<div align="center">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Դ</div></td>
            <td><input style="width:100%;" type="text" name="Source" value="<% = SourceStr %>"></td>
          </tr>
          <tr> 
            <td height="26"> 
              <div align="center">�ɼ�����</div></td>
            <td><input name="AddDate" type="text" id="AddDate2" style="width:100%;" value="<% = AddDate %>"> 
              <div align="center"></div></td>
          </tr>
        </table></td>
	</tr>
    <tr> 
      <td height="20" colspan="2">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td id="EditCodeBtn" width="100" class="LableSelected" onClick="CodeContent();" bgcolor="#EFEFEF"> <div align="center"> 
                �� ��</div></td>
            <td width="5" class="ToolBarButtonLine">&nbsp;</td>
			<td id="PreviewBtn" width="100" class="LableDefault" onClick="Preview();"> <div align="center">Ԥ 
                ��</div></td>
            <td class="ToolBarButtonLine">&nbsp;</td>
          </tr>
        </table></td>
    </tr>
    <tr id="EditCodeArea" bgcolor="#EFEFEF"> 
      <td height="300" colspan="2" class="LableWindow"> 
        <textarea name="Content" id="NewsContent" rows="20" style="width:100%;"><% = Content %></textarea></td>
    </tr>
    <tr id="PreviewArea" style="display:none;" bgcolor="#EFEFEF"> 
      <td height="300" colspan="2" class="LableWindow"> 
        <iframe name="PreviewContent" frameborder="1" class="Composition" ID="PreviewContent" MARGINHEIGHT="1" MARGINWIDTH="1" width="100%" scrolling="yes" src="about:blank"></iframe></td>
    </tr>
</table>
</form>
</body>
</html>
<%
Set CollectConn = Nothing
Set Conn = Nothing
Set RsNewsObj = Nothing
%>
<script language="JavaScript">
function SetEditAreaHeight()
{
	var BodyHeight=document.body.clientHeight;
	var EditAreaHeight=BodyHeight-200;
	document.all.NewsContent.style.height=EditAreaHeight;
	document.all.PreviewContent.height=EditAreaHeight;
}
SetEditAreaHeight();
window.onresize=SetEditAreaHeight;
function Preview()
{
	var TempStr='';
	document.all.EditCodeArea.style.display='none';
	document.all.PreviewArea.style.display='';
	PreviewContent.document.write('<head><link href=\"../../CSS/FS_css.css\" type=\"text/css\" rel=\"stylesheet\"></head><body MONOSPACE>');
	PreviewContent.document.body.innerHTML=document.all.Content.value;
	document.all.PreviewBtn.className='LableSelected';
	document.all.PreviewBtn.style.backgroundColor='#EFEFEF';
	document.all.EditCodeBtn.className='LableDefault';
	document.all.EditCodeBtn.style.backgroundColor='#FFFFFF';
}
function CodeContent()
{
	document.all.EditCodeArea.style.display='';
	document.all.PreviewArea.style.display='none';
	document.all.EditCodeBtn.className='LableSelected';
	document.all.EditCodeBtn.style.backgroundColor='#EFEFEF';
	document.all.PreviewBtn.className='LableDefault';
	document.all.PreviewBtn.style.backgroundColor='#FFFFFF';
}
</script>