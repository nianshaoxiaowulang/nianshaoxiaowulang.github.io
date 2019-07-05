<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="inc/Config.asp" -->
<!--#include file="inc/Function.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<%
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
Dim DBC,Conn,CollectConn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
DBC.ConnStr = CollectDBConnectionStr
Set CollectConn = DBC.OpenConnection()
Set DBC = Nothing
'判断权限
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P080105") then Call ReturnError1()
'判断权限结束
Dim RsEditObj,EditSql,SiteID
Dim ObjUrl
Set RsEditObj = Server.CreateObject ("ADODB.RecordSet")
SiteID = Request("SiteID")
if SiteID <> "" then
	EditSql="Select * from FS_Site where ID=" & SiteID
	RsEditObj.Open EditSql,CollectConn,1,3
	if RsEditObj.Eof then
		Response.write"<script>alert(""没有修改的站点"");location.href=""javascript:history.back()"";</script>"
		Response.end
	else
		ObjUrl = RsEditObj("ObjUrl")
	end if
else
	Response.write"<script>alert(""没有修改的站点"");location.href=""javascript:history.back()"";</script>"
	Response.end
end if

On Error Resume Next 
Dim ListHeadSetting,ListFootSetting,OtherPageHeadSetting,OtherPageFootSetting
Dim IndexRule,StartPageNum,EndPageNum,HandPageContent,OtherType
Dim ListSetting,OtherPageSetting
ListSetting = split(Request.Form("ListSetting"),"[列表内容]",-1,1)
ListHeadSetting = ListSetting(0)
ListFootSetting = ListSetting(1)
If Err Or ListHeadSetting="" Or ListFootSetting="" Then
	ListHeadSetting = "<body"
	ListFootSetting = "</body>"
	Err.clear
End If
If InStr(Request.Form("OtherPageSetting"),"[其他页面]")<>0 then
	OtherPageSetting = split(Request.Form("OtherPageSetting"),"[其他页面]",-1,1)
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
<title>自动新闻采集—站点设置</title>
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
			  <td width="50" align="center" alt="第三步" onClick="window.location.href='javascript:history.go(-1)';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">上一步</td>
			  <td width=2 class="Gray">|</td>
			  <td width="50" align="center" alt="第四步" onClick="document.form1.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">下一步</td>
			  <td width=2 class="Gray">|</td>
		      <td width="35" align="center" alt="后退" onClick="history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
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
        <div align="center">列表URL</div></td>
		<td>	&nbsp;&nbsp;输入区域：
			<span onClick="if(document.Form1.LinkSetting.rows>2)document.Form1.LinkSetting.rows-=1" style='cursor:hand'><b>缩小</b></span>
			<span onClick="document.Form1.LinkSetting.rows+=1" style='cursor:hand'><b>扩大</b></span>
		&nbsp;&nbsp;可用标签：<font onClick="addTag('[列表URL]')" style="CURSOR: hand"><b>[列表URL]</b></font>&nbsp;&nbsp;&nbsp;<font onClick="addTag('[变量]')" style="CURSOR: hand"><b>[变量]</b></font><br>
		 <textarea onfocus="getActiveText(this)" onclick="getActiveText(this)"  onchange="getActiveText(this)" name="LinkSetting" cols="50" rows="6" id="textarea2" style="width:100%;"><%=RsEditObj("LinkHeadSetting")%>[列表URL]<%=RsEditObj("LinkFootSetting")%></textarea></td>
	  </tr>
</table>
</form>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="28" class="ButtonListLeft"> 
      <div align="center">代码</div></td>
  </tr>
  <tr>
    <td height="20"><textarea name="CodeArea" rows="18" style="width:100%;"></textarea></td>
  </tr>
  <tr> 
    <td height="28" class="ButtonListLeft"> 
      <div align="center">结果</div></td>
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