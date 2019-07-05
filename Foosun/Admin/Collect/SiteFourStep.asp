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
Dim LinkHeadSetting,LinkFootSetting
Dim ObjUrl,ListHeadSetting,ListFootSetting,NewsLinkStr
Dim HandSetAuthor,HandSetSource,HandSetAddDate
Set RsEditObj = Server.CreateObject ("ADODB.RecordSet")
SiteID = Request("SiteID")
if SiteID <> "" then
	EditSql="Select * from FS_Site where ID=" & SiteID
	RsEditObj.Open EditSql,CollectConn,1,3
	if RsEditObj.Eof then
		Response.write("没有修改的站点")
	else
		ObjUrl = RsEditObj("ObjUrl")
		ListHeadSetting = RsEditObj("ListHeadSetting")
		ListFootSetting = RsEditObj("ListFootSetting")
		HandSetAuthor = RsEditObj("HandSetAuthor")
		HandSetSource = RsEditObj("HandSetSource")
		HandSetAddDate = RsEditObj("HandSetAddDate")
	end if
else
	Response.write("没有修改的站点")
end if
Dim ListSetting
If InStr(Request.Form("LinkSetting"),"[列表URL]") = 0 Then
	Response.Write "<script>alert('列表URL没有设置或设置不正确！');history.back();</script>"
	Response.End 
End if
ListSetting = Split(Request.Form("LinkSetting"),"[列表URL]",-1,1)
LinkHeadSetting = ListSetting(0)
LinkFootSetting = ListSetting(1)

if Request.Form("Result") = "Edit" then
    Dim RsAddObj,sql
	Set RsAddObj = Server.CreateObject ("ADODB.RecordSet")
	Sql = "select * from FS_Site where id=" & Request.Form("SiteID")
	RsAddObj.Open Sql,CollectConn,1,3
	RsAddObj("LinkHeadSetting") = LinkHeadSetting
	RsAddObj("LinkFootSetting") = LinkFootSetting
	RsAddObj.update
	RsAddObj.close
	Set RsAddObj = Nothing
end if

Dim ResponseAllStr,NewsListStr
ResponseAllStr = GetPageContent(ObjURL)
NewsListStr = GetOtherContent(ResponseAllStr,ListHeadSetting,ListFootSetting)
NewsLinkStr = FormatUrl(GetOtherContent(NewsListStr,LinkHeadSetting,LinkFootSetting),ObjUrl)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>自动新闻采集—站点设置</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<body leftmargin="2" topmargin="2">
<form name="form1" method="post" action="SiteFiveStep.asp" id="Form1">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td width="50" align="center" alt="第四步" onClick="window.location.href='javascript:history.go(-1)';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">上一步</td>
			<td width=2 class="Gray">|</td>
            <td width="50" align="center" alt="第五步" onClick="document.form1.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">下一步</td>
			<td width=2 class="Gray">|</td>
		    <td width="35" align="center" alt="后退" onClick="history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
            <td>&nbsp; <input name="SiteID" type="hidden" id="SiteID2" value="<% = SiteID %>"> 
              <input name="Result" type="hidden" id="Result2" value="Edit"> <input type="hidden" name="NewsLinkStr" value="<% = NewsLinkStr %>"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
    <tr> 
      <td width="20%"> <div align="center">标题</div></td>
      <td>	&nbsp;&nbsp;输入区域：
			<span onClick="if(document.Form1.PageTitleSetting.rows>2)document.Form1.PageTitleSetting.rows-=1" style='cursor:hand'><b>缩小</b></span>
			<span onClick="document.Form1.PageTitleSetting.rows+=1" style='cursor:hand'><b>扩大</b></span>
	  &nbsp;&nbsp;可用标签：<font onmouseover="getActiveText(document.form1.PageTitleSetting);" onClick="addTag('[标题]')" style="CURSOR: hand"><b>[标题]</b></font>&nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.form1.PageTitleSetting);" onClick="addTag('[变量]')" style="CURSOR: hand"><b>[变量]</b></font>
        <table width="95%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td height="5"></td>
          </tr>
        </table>
        <textarea ondblclick="getActiveText(this)" onclick="getActiveText(this)"  onchange="getActiveText(this)" name="PageTitleSetting" cols="50" rows="3" id="textarea4" style="width:100%;"><%=RsEditObj("PageTitleHeadSetting")%>[标题]<%=RsEditObj("PageTitleFootSetting")%></textarea></td>
    </tr>
    <tr> 
      <td> <div align="center">内容</div></td>
      <td> &nbsp;&nbsp;输入区域： <span onClick="if(document.Form1.PagebodySetting.rows>2)document.Form1.PagebodySetting.rows-=1" style='cursor:hand'><b>缩小</b></span> 
        <span onClick="document.Form1.PagebodySetting.rows+=1" style='cursor:hand'><b>扩大</b></span> 
        &nbsp;&nbsp;可用标签：<font onmouseover="getActiveText(document.form1.PagebodySetting);" onClick="addTag('[内容]')" style="CURSOR: hand"><b>[内容]</b></font>&nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.form1.PagebodySetting);" onClick="addTag('[变量]')" style="CURSOR: hand"><b>[变量]</b></font><textarea onDblClick="getActiveText(this)" onClick="getActiveText(this)"  onChange="getActiveText(this)" name="PagebodySetting" cols="50" rows="3" id="textarea" style="width:100%;"><%=RsEditObj("PagebodyHeadSetting")%>[内容]<%=RsEditObj("PagebodyFootSetting")%></textarea></td>
    </tr>
    <tr> 
      <td height="26" colspan="4"> <div align="left"> 　　　　　　　　　　　　　　　　　
<input name="OtherSetType" type="radio" onClick="ChangeSetOption(0);" value="0" checked>
          设置作者 
          <input type="radio" name="OtherSetType" onClick="ChangeSetOption(1);" value="1">
          设置来源 
          <input type="radio" name="OtherSetType" onClick="ChangeSetOption(2);" value="2">
          设置时间 
          <input type="radio" name="OtherSetType" onClick="ChangeSetOption(3);" value="3">
          设置分页 
        </div></td>
    </tr>
    <tr id="SetAuthor" style="display:;"> 
      <td colspan="4"><table width="100%" border="0" cellspacing="0" cellpadding="3">
          <tr>
            <td height="26">
<div align="center">手动设置</div></td>
            <td colspan="3"><input style="width:100%;" type="text" name="HandSetAuthor" value="<% = HandSetAuthor %>"></td>
          </tr>
          <tr> 
            <td width="20%"> <div align="center">作者</div></td>
            <td colspan="3">&nbsp;&nbsp;输入区域：
			<span onClick="if(document.Form1.AuthorSetting.rows>2)document.Form1.AuthorSetting.rows-=1" style='cursor:hand'><b>缩小</b></span>
			<span onClick="document.Form1.AuthorSetting.rows+=1" style='cursor:hand'><b>扩大</b></span>
			 &nbsp;&nbsp;可用标签：<font onmouseover="getActiveText(document.form1.AuthorSetting);" onClick="addTag('[作者]')" style="CURSOR: hand"><b>[作者]</b></font>&nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.form1.AuthorSetting);" onClick="addTag('[变量]')" style="CURSOR: hand"><b>[变量]</b></font>
              <table width="95%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="5"></td>
                </tr>
              </table>
              <textarea ondblclick="getActiveText(this)" onclick="getActiveText(this)"  onchange="getActiveText(this)" name="AuthorSetting" cols="50" rows="3" id="textarea9" style="width:100%;"><%=RsEditObj("AuthorHeadSetting")%>[作者]<%=RsEditObj("AuthorFootSetting")%></textarea></td>
          </tr>
        </table></td>
    </tr>
    <tr id="SetSource" style="display:none;"> 
      <td colspan="4"><table width="100%" border="0" cellspacing="0" cellpadding="3">
          <tr>
            <td height="26">
<div align="center">手动设置</div></td>
            <td colspan="3"><input style="width:100%;" type="text" name="HandSetSource" value="<% = HandSetSource %>"></td>
          </tr>
		  <tr> 
            <td width="20%"> <div align="center">来源</div></td>
            <td colspan="3">&nbsp;&nbsp;输入区域：
			<span onClick="if(document.Form1.SourceSetting.rows>2)document.Form1.SourceSetting.rows-=1" style='cursor:hand'><b>缩小</b></span>
			<span onClick="document.Form1.SourceSetting.rows+=1" style='cursor:hand'><b>扩大</b></span>
			&nbsp;&nbsp;可用标签：<font onmouseover="getActiveText(document.form1.SourceSetting);" onClick="addTag('[来源]')" style="CURSOR: hand"><b>[来源]</b></font>&nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.form1.SourceSetting);" onClick="addTag('[变量]')" style="CURSOR: hand"><b>[变量]</b></font>
              <table width="95%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="5"></td>
                </tr>
              </table>
              <textarea ondblclick="getActiveText(this)" onclick="getActiveText(this)"  onchange="getActiveText(this)" name="SourceSetting" cols="50" rows="3" id="textarea9a" style="width:100%;"><%=RsEditObj("SourceHeadSetting")%>[来源]<%=RsEditObj("SourceFootSetting")%></textarea></td>
          </tr>
        </table></td>
    </tr>
    <tr id="SetAddTime" style="display:none;"> 
      <td colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="3">
          <tr>
            <td height="26">
<div align="center">手动设置</div></td>
            <td colspan="3"><input style="width:100%;" type="text" name="HandSetAddDate" value="<% = HandSetAddDate %>"></td>
          </tr>
		  <tr> 
            <td width="20%"> <div align="center">加入时间</div></td>
            <td>&nbsp;&nbsp;输入区域：
			<span onClick="if(document.Form1.AddDateSetting.rows>2)document.Form1.AddDateSetting.rows-=1" style='cursor:hand'><b>缩小</b></span>
			<span onClick="document.Form1.AddDateSetting.rows+=1" style='cursor:hand'><b>扩大</b></span>
			&nbsp;&nbsp;可用标签：<font onmouseover="getActiveText(document.form1.AddDateSetting);" onClick="addTag('[加入时间]')" style="CURSOR: hand"><b>[加入时间]</b></font>&nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.form1.AddDateSetting);" onClick="addTag('[变量]')" style="CURSOR: hand"><b>[变量]</b></font>
              <table width="95%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="5"></td>
                </tr>
              </table>
              <textarea ondblclick="getActiveText(this)" onclick="getActiveText(this)"  onchange="getActiveText(this)" name="AddDateSetting" cols="50" rows="3" id="textarea9" style="width:100%;"><%=RsEditObj("AddDateHeadSetting")%>[加入时间]<%=RsEditObj("AddDateFootSetting")%></textarea></td>
          </tr>
        </table></td>
    </tr>
    <tr id="SetCutPage" style="display:none;"> 
      <td colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="3">
          <tr> 
            <td width="20%"> 
              <div align="center">分页新闻<br>(下一页)</div></td>
      <td>&nbsp;&nbsp;输入区域：
			<span onClick="if(document.Form1.OtherNewsPageSetting.rows>2)document.Form1.OtherNewsPageSetting.rows-=1" style='cursor:hand'><b>缩小</b></span>
			<span onClick="document.Form1.OtherNewsPageSetting.rows+=1" style='cursor:hand'><b>扩大</b></span>
	  &nbsp;&nbsp;可用标签：<font onmouseover="getActiveText(document.form1.OtherNewsPageSetting);" onClick="addTag('[分页新闻]')" style="CURSOR: hand"><b>[分页新闻]</b></font>&nbsp;&nbsp;&nbsp;<font onmouseover="getActiveText(document.form1.OtherNewsPageSetting);" onClick="addTag('[变量]')" style="CURSOR: hand"><b>[变量]</b></font>
              <table width="95%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="5"></td>
                </tr>
              </table>
              <textarea ondblclick="getActiveText(this)" onclick="getActiveText(this)"  onchange="getActiveText(this)" name="OtherNewsPageSetting" cols="50" rows="3" id="textarea5" style="width:100%;"><%=RsEditObj("OtherNewsPageHeadSetting")%>[分页新闻]<%=RsEditObj("OtherNewsPageFootSetting")%></textarea></td>
    </tr>
        </table></td>
    </tr>
</table>
</form>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td colspan="2" height="28" class="ButtonListLeft"> 
      <div align="center">预览结果</div></td>
  </tr>
  <tr> 
    <td height="36" colspan="2">
<div align="center"><a href="<% = NewsLinkStr %>" target="_blank"> 
        <% = NewsLinkStr %>
        </a></div></td>
  </tr>
</table>
</body>
</html>
<%
Set RsEditObj = Nothing
Set CollectConn = Nothing
%>
<script language="JavaScript">
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