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
Dim NewsLinkStr
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

Dim PageTitleHeadSetting,PageTitleFootSetting,PagebodyHeadSetting,PagebodyFootSetting
Dim OtherNewsPageHeadSetting,OtherNewsPageFootSetting
Dim AuthorHeadSetting,AuthorFootSetting
Dim SourceHeadSetting,SourceFootSetting
Dim AddDateHeadSetting,AddDateFootSetting
Dim HandSetAuthor,HandSetSource,HandSetAddDate
Dim TextTF,IsStyle,IsDiv,IsA,IsClass,IsFont,IsSpan,IsObjectTF,IsIFrame,IsScript
Dim PageTitleSetting,PagebodySetting,OtherNewsPageSetting,AuthorSetting,SourceSetting,AddDateSetting
If InStr(Request.Form("PageTitleSetting"),"[����]") = 0 Then
	Response.Write "<script>alert('���ű���û�����û����ò���ȷ��');history.back();</script>"
	Response.End 
End If
If InStr(Request.Form("PagebodySetting"),"[����]") = 0 Then
	Response.Write "<script>alert('��������û�����û����ò���ȷ��');history.back();</script>"
	Response.End 
End if
PageTitleSetting = Split(Request.Form("PageTitleSetting"),"[����]",-1,1)
PageTitleHeadSetting = PageTitleSetting(0)
PageTitleFootSetting = PageTitleSetting(1)
PagebodySetting = Split(Request.Form("PagebodySetting"),"[����]",-1,1)
PagebodyHeadSetting = PagebodySetting(0)
PagebodyFootSetting = PagebodySetting(1)
If InStr(Request.Form("OtherNewsPageSetting"),"[��ҳ����]")<>0 then
	OtherNewsPageSetting = Split(Request.Form("OtherNewsPageSetting"),"[��ҳ����]",-1,1)
	OtherNewsPageHeadSetting = OtherNewsPageSetting(0)
	OtherNewsPageFootSetting = OtherNewsPageSetting(1)
End If
If InStr(Request.Form("AuthorSetting"),"[����]")<>0 then
	AuthorSetting = Split(Request.Form("AuthorSetting"),"[����]",-1,1)
	AuthorHeadSetting = AuthorSetting(0)
	AuthorFootSetting = AuthorSetting(1)
End If 
If InStr(Request.Form("SourceSetting"),"[��Դ]")<>0 then
	SourceSetting = Split(Request.Form("SourceSetting"),"[��Դ]",-1,1)
	SourceHeadSetting = SourceSetting(0)
	SourceFootSetting = SourceSetting(1)
End If
If InStr(Request.Form("AddDateSetting"),"[����ʱ��]")<>0 then
	AddDateSetting = Split(Request.Form("AddDateSetting"),"[����ʱ��]",-1,1)
	AddDateHeadSetting = AddDateSetting(0)
	AddDateFootSetting = AddDateSetting(1)
End If 
HandSetAuthor = Request.Form("HandSetAuthor")
HandSetSource = Request.Form("HandSetSource")
HandSetAddDate = Request.Form("HandSetAddDate")
if Request.Form("Result") = "Edit" then
    Dim RsAddObj,sql
	Set RsAddObj = Server.CreateObject ("ADODB.RecordSet")
	Sql = "select * from FS_Site where id=" & Request.Form("SiteID")
	RsAddObj.Open Sql,CollectConn,1,3
	TextTF = RsAddObj("TextTF")
	IsStyle = RsAddObj("IsStyle")
	IsDiv = RsAddObj("IsDiv")
	IsA = RsAddObj("IsA")
	IsClass = RsAddObj("IsClass")
	IsFont = RsAddObj("IsFont")
	IsSpan = RsAddObj("IsSpan")
	IsObjectTF = RsAddObj("IsObject")
	IsIFrame = RsAddObj("IsIFrame")
	IsScript = RsAddObj("IsScript")
	RsAddObj("PagebodyHeadSetting") = PagebodyHeadSetting
	RsAddObj("PagebodyFootSetting") = PagebodyFootSetting
	RsAddObj("PageTitleHeadSetting") = PageTitleHeadSetting
	RsAddObj("PageTitleFootSetting") = PageTitleFootSetting
	RsAddObj("OtherNewsPageHeadSetting") = OtherNewsPageHeadSetting
	RsAddObj("OtherNewsPageFootSetting") = OtherNewsPageFootSetting
	RsAddObj("AuthorHeadSetting") = AuthorHeadSetting
	RsAddObj("AuthorFootSetting") = AuthorFootSetting
	RsAddObj("SourceHeadSetting") = SourceHeadSetting
	RsAddObj("SourceFootSetting") = SourceFootSetting
	RsAddObj("AddDateHeadSetting") = AddDateHeadSetting
	RsAddObj("AddDateFootSetting") = AddDateFootSetting
	RsAddObj("HandSetAuthor") = HandSetAuthor
	RsAddObj("HandSetSource") = HandSetSource
	if IsDate(HandSetAddDate) then
		RsAddObj("HandSetAddDate") = HandSetAddDate
	end if
	RsAddObj.update
	RsAddObj.close
	Set RsAddObj = Nothing
end if

NewsLinkStr = Request("NewsLinkStr")
Dim ResponseAllStr,TitleStr,NewsBodyStr,AuthorStr,SourceStr,AddDateStr
ResponseAllStr = GetPageContent(NewsLinkStr)
TitleStr = GetOtherContent(ResponseAllStr,PageTitleHeadSetting,PageTitleFootSetting)
NewsBodyStr = GetOtherContent(ResponseAllStr,PagebodyHeadSetting,PagebodyFootSetting)
NewsBodyStr = ReplaceContentStr(NewsBodyStr)
if HandSetAuthor <> "" then
	AuthorStr = HandSetAuthor
else
	if AuthorHeadSetting <> "" And AuthorFootSetting <> "" then 
		AuthorStr = GetOtherContent(ResponseAllStr,AuthorHeadSetting,AuthorFootSetting)
	end if
end if
if HandSetSource <> "" then
	SourceStr = HandSetSource
else
	if SourceHeadSetting <> "" And SourceFootSetting <> "" then 
		SourceStr = GetOtherContent(ResponseAllStr,SourceHeadSetting,SourceFootSetting)
	end if
end if
if HandSetAddDate <> "" then
	if Not IsDate(HandSetAddDate) then
		AddDateStr = Now
	else
		AddDateStr = HandSetAddDate
	end if
else
	if AddDateHeadSetting <> "" And AddDateFootSetting <> "" then 
		AddDateStr = GetOtherContent(ResponseAllStr,AddDateHeadSetting,AddDateFootSetting)
	end if
end if
NewsBodyStr = Replace(Replace(NewsBodyStr,"""","%22"),"'","%27")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�Զ����Ųɼ���վ������</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<body leftmargin="2" topmargin="2">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td width="50" align="center" alt="���Ĳ�" onClick="window.location.href='javascript:history.go(-1)';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">��һ��</td>
			<td width=2 class="Gray">|</td>
            <td width="35" align="center" alt="���" onClick="window.location.href='Site.asp';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">���</td>
			<td width=2 class="Gray">|</td>
		    <td width="35" align="center" alt="����" onClick="history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
            <td><input type="hidden" name="NewsBodyStr" value="<% = NewsBodyStr %>"> &nbsp;</td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="26">
<div align="center"><strong><font size="3"><% = TitleStr %></font></strong></div></td>
	</tr>
	<tr>
	  
    <td height="26">
<div align="center"><strong>����</strong>�� 
        <% = AuthorStr %>
        &nbsp;&nbsp;<strong>��Դ</strong>�� 
        <% = SourceStr %>
        &nbsp;&nbsp;<strong>ʱ��</strong>�� 
        <% = AddDateStr %></div></td>
	</tr>
	<tr>
	  <td><iframe frameborder="1" name="PreviewArea" src="about:blank" ID="PreviewArea" MARGINHEIGHT="1" MARGINWIDTH="1" height="480" width="100%" scrolling="yes"></iframe></td>
	</tr>
</table>
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
	frames["PreviewArea"].document.write(unescape(document.all.NewsBodyStr.value));
}
</script>