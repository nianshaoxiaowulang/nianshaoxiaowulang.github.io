<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="inc/Config.asp" -->
<!--#include file="inc/Function.asp" -->
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
if Not JudgePopedomTF(Session("Name"),"P080105") then Call ReturnError1()
'�ж�Ȩ�޽���
Dim RsEditObj,EditSql,SiteID
Dim ObjUrl
Set RsEditObj = Server.CreateObject ("ADODB.RecordSet")
SiteID = Request("SiteID")
if SiteID <> "" then
	EditSql="Select * from FS_Site where ID=" & SiteID
	RsEditObj.Open EditSql,CollectConn,1,3
	if RsEditObj.Eof then
		Response.write"<script>alert(""û���޸ĵ�վ��"");location.href=""javascript:history.back()"";</script>"
		Response.end
	else
		ObjUrl = RsEditObj("ObjUrl")
	end if
else
	Response.write"<script>alert(""û���޸ĵ�վ��"");location.href=""javascript:history.back()"";</script>"
	Response.end
end if

On Error Resume Next 
Dim ListHeadSetting,ListFootSetting,OtherPageHeadSetting,OtherPageFootSetting
Dim IndexRule,StartPageNum,EndPageNum,HandPageContent,OtherType
Dim ListSetting,OtherPageSetting
ListSetting = split(Request.Form("ListSetting"),"[�б�����]",-1,1)
ListHeadSetting = ListSetting(0)
ListFootSetting = ListSetting(1)
If Err Or ListHeadSetting="" Or ListFootSetting="" Then
	ListHeadSetting = "<body"
	ListFootSetting = "</body>"
	Err.clear
End If
If InStr(Request.Form("OtherPageSetting"),"[����ҳ��]")<>0 then
	OtherPageSetting = split(Request.Form("OtherPageSetting"),"[����ҳ��]",-1,1)
	OtherPageHeadSetting = OtherPageSetting(0)
	OtherPageFootSetting = OtherPageSetting(1)
End if
OtherType = Request.Form("OtherType")
IndexRule = Request.Form("IndexRule")
StartPageNum = Request.Form("StartPageNum")
EndPageNum = Request.Form("EndPageNum")
HandPageContent = Request.Form("HandPageContent")
if Request.Form("Result") = "Edit" then
    Dim RsAddObj,sql
	Set RsAddObj = Server.CreateObject ("ADODB.RecordSet")
	Sql = "select * from FS_Site where id=" & Request.Form("SiteID")
	RsAddObj.Open Sql,CollectConn,1,3
	RsAddObj("ListHeadSetting") = ListHeadSetting
	RsAddObj("ListFootSetting") = ListFootSetting
	Select Case OtherType
		Case "0"
			RsAddObj("OtherType") = OtherType
			RsAddObj("IndexRule") = ""
			RsAddObj("StartPageNum") = ""
			RsAddObj("EndPageNum") = ""
			RsAddObj("HandPageContent") = ""
			RsAddObj("OtherPageHeadSetting") = ""
			RsAddObj("OtherPageFootSetting") = ""
		Case "1"
			RsAddObj("OtherType") = OtherType
			RsAddObj("IndexRule") = ""
			RsAddObj("StartPageNum") = ""
			RsAddObj("EndPageNum") = ""
			RsAddObj("HandPageContent") = ""
			RsAddObj("OtherPageHeadSetting") = OtherPageHeadSetting
			RsAddObj("OtherPageFootSetting") = OtherPageFootSetting
		Case "2"
			RsAddObj("OtherType") = OtherType
			RsAddObj("IndexRule") = IndexRule
			RsAddObj("StartPageNum") = StartPageNum
			RsAddObj("EndPageNum") = EndPageNum
			RsAddObj("HandPageContent") = ""
			RsAddObj("OtherPageHeadSetting") = ""
			RsAddObj("OtherPageFootSetting") = ""
		Case "3"
			RsAddObj("OtherType") = OtherType
			RsAddObj("IndexRule") = ""
			RsAddObj("StartPageNum") = ""
			RsAddObj("EndPageNum") = ""
			RsAddObj("HandPageContent") = HandPageContent
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
	RsAddObj.update
	RsAddObj.close
	Set RsAddObj = Nothing
end if
Dim ResponseAllStr,NewsListStr
ResponseAllStr = GetPageContent(ObjURL)
NewsListStr = GetOtherContent(ResponseAllStr,ListHeadSetting,ListFootSetting)
NewsListStr = Replace(Replace(NewsListStr,"""","%22"),"'","%27")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�Զ����Ųɼ���վ������</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<body leftmargin="2" topmargin="2">
<form name="form1" method="post" action="SiteFourStep.asp" id="Form1">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
			  <td width="50" align="center" alt="������" onClick="window.location.href='javascript:history.go(-1)';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">��һ��</td>
			  <td width=2 class="Gray">|</td>
			  <td width="50" align="center" alt="���Ĳ�" onClick="document.form1.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">��һ��</td>
			  <td width=2 class="Gray">|</td>
		      <td width="35" align="center" alt="����" onClick="history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
			  <td>&nbsp; <input name="SiteID" type="hidden" id="SiteID2" value="<% = SiteID %>"> 
				<input name="Result" type="hidden" id="Result2" value="Edit">
              <input type="hidden" name="NewsListStr" value="<% = NewsListStr %>"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	  <tr> 
      <td width="20%"> 
        <div align="center">�б�URL</div></td>
		<td>	&nbsp;&nbsp;��������
			<span onClick="if(document.Form1.LinkSetting.rows>2)document.Form1.LinkSetting.rows-=1" style='cursor:hand'><b>��С</b></span>
			<span onClick="document.Form1.LinkSetting.rows+=1" style='cursor:hand'><b>����</b></span>
		&nbsp;&nbsp;���ñ�ǩ��<font onClick="addTag('[�б�URL]')" style="CURSOR: hand"><b>[�б�URL]</b></font>&nbsp;&nbsp;&nbsp;<font onClick="addTag('[����]')" style="CURSOR: hand"><b>[����]</b></font><br>
		 <textarea onfocus="getActiveText(this)" onclick="getActiveText(this)"  onchange="getActiveText(this)" name="LinkSetting" cols="50" rows="6" id="textarea2" style="width:100%;"><%=RsEditObj("LinkHeadSetting")%>[�б�URL]<%=RsEditObj("LinkFootSetting")%></textarea></td>
	  </tr>
</table>
</form>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="28" class="ButtonListLeft"> 
      <div align="center">����</div></td>
  </tr>
  <tr>
    <td height="20"><textarea name="CodeArea" rows="18" style="width:100%;"></textarea></td>
  </tr>
  <tr> 
    <td height="28" class="ButtonListLeft"> 
      <div align="center">���</div></td>
  </tr>
  <tr> 
    <td><iframe frameborder="1" name="PreviewArea" src="about:blank" ID="PreviewArea" MARGINHEIGHT="1" MARGINWIDTH="1" height="300" width="100%" scrolling="yes"></iframe></td>
  </tr>
</table>
<p><p><p>
</body>
</html>
<%
Set CollectConn = Nothing
Set Conn = Nothing
Set RsEditObj = Nothing
%>
<script language="JavaScript">
function document.onreadystatechange()
{
	document.all.CodeArea.value=unescape(document.form1.NewsListStr.value);
	frames["PreviewArea"].document.write(unescape(document.form1.NewsListStr.value));
}

currObj = "uuuu";
function getActiveText(obj)
{
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