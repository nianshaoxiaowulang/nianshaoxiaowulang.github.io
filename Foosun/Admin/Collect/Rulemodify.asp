<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="inc/Config.asp" -->
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
if Not JudgePopedomTF(Session("Name"),"P080202") then Call ReturnError1()
'判断权限结束
Dim RuleID
RuleID = Request("RuleID")
if Request.Form("Result")="Edit" then
    Dim Sql,RsEditObj
	if RuleID <> "" then
		Set RsEditObj = Server.CreateObject ("ADODB.RecordSet")
		Sql = "Select * from FS_Rule where id=" & RuleID
		RsEditObj.Open Sql,CollectConn,1,3
		if RsEditObj.Eof then
			Response.Write"<script>alert(""没有修改规则"");location.href=""javascript:history.back()"";</script>"
			Response.End
		end if
		RsEditObj("RuleName") = NoCSSHackAdmin(Request.Form("RuleName"),"规则名称")
		RsEditObj("SiteId") = Request.Form("SiteId")
		Dim KeywordSetting
		If InStr(Request.Form("KeywordSetting"),"[过滤字符串]")<>0 then
			KeywordSetting = Split(Request.Form("KeywordSetting"),"[过滤字符串]",-1,1)
			RsEditObj("HeadSeting") = KeywordSetting(0)
			RsEditObj("FootSeting") = KeywordSetting(1)
		End If
		RsEditObj("ReContent") = Request.Form("ReContent")
		RsEditObj.UpDate
		RsEditObj.Close
		Set RsEditObj = Nothing
	else
		Response.Write"<script>alert(""参数传递错误"");location.href=""javascript:history.back()"";</script>"
		Response.End
	end if
	Response.Redirect("Rule.asp")
	Response.End
end if

Dim RsRuleObj
if RuleID <> "" then
	Set RsRuleObj = Server.CreateObject ("ADODB.RecordSet")
	Sql = "Select * from FS_Rule where id=" & RuleID
	RsRuleObj.Open Sql,CollectConn,1,3
	if RsRuleObj.Eof then
		Response.Write"<script>alert(""没有修改规则"");location.href=""javascript:history.back()"";</script>"
		Response.End
	end if
else
	Response.Write"<script>alert(""参数传递错误"");location.href=""javascript:history.back()"";</script>"
	Response.End
end if
	
Dim SiteList,RsSiteObj
Set RsSiteObj = Server.CreateObject("Adodb.RecordSet")
RsSiteObj.Source = "Select ID,SiteName from FS_Site order by id desc"
RsSiteObj.open RsSiteObj.Source,CollectConn,1,3
do while Not RsSiteObj.Eof
	if Clng(RsRuleObj("SiteID")) = Clng(RsSiteObj("ID")) then
		SiteList = SiteList & "<option selected value=" & RsSiteObj("ID") & "" & ">" & RsSiteObj("SiteName") & "</option><br>"
	else
		SiteList = SiteList & "<option value=" & RsSiteObj("ID") & "" & ">" & RsSiteObj("SiteName") & "</option><br>"
	end if
	RsSiteObj.MoveNext	
loop
RsSiteObj.Close
Set RsSiteObj = Nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>自动新闻采集—站点设置</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<body leftmargin="2" topmargin="2">
<form name="form1" id="form1" method="post" action="">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td width="35" align="center" alt="保存" onClick="document.form1.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td>
			<td width=2 class="Gray">|</td>
		    <td width="35" align="center" alt="后退" onClick="history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
            <td>&nbsp; <input name="Result" type="hidden" id="Result4" value="Edit">
          <input name="id" type="hidden" id="id2" value="<% = RuleID %>"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="100%" border="0" cellpadding="3" cellspacing="1">
    <tr> 
      <td width="100"> <div align="center">规则名称</div></td>
      <td> <input name="RuleName" style="width:100%;" type="text" id="RuleName" value="<% = RsRuleObj("RuleName") %>"> 
        <div align="right"></div></td>
    </tr>
    <tr> 
      <td><div align="center">应用到</div></td>
      <td><select name="SiteId" style="width:100%;" id="SiteId">
          <% =SiteList %>
        </select></td>
    </tr>
    <tr> 
      <td> <div align="center">过滤字符串</div></td>
      <td> &nbsp;&nbsp;输入区域： <span onClick="if(document.Form1.KeywordSetting.rows>2)document.Form1.KeywordSetting.rows-=1" style='cursor:hand'><b>缩小</b></span> 
        <span onClick="document.Form1.KeywordSetting.rows+=1" style='cursor:hand'><b>扩大</b></span> 
        &nbsp;&nbsp;可用标签:<font onClick="addTag('[过滤字符串]')" style="CURSOR: hand"><b>[过滤字符串]</b></font> 
        &nbsp;&nbsp;&nbsp;<font onClick="addTag('[变量]')" style="CURSOR: hand"><b>[变量]</b></font><br>
        <br>
	  <textarea name="KeywordSetting"  onfocus="getActiveText(this)" onclick="getActiveText(this)"  onchange="getActiveText(this)" rows="5" id="textarea2" style="width:100%;"><% = RsRuleObj("HeadSeting") %>[过滤字符串]<% = RsRuleObj("FootSeting") %></textarea> 
	  </td>
    </tr>
    <tr> 
      <td> <div align="center"> 
          替换为</div></td>
      <td colspan="3"><textarea style="width:100%;" name="ReContent" cols="30" rows="5" id="ReContent"><% = RsRuleObj("ReContent") %></textarea></td>
    </tr>
</table>
</form>
</body>
</html>
<%
Set CollectConn = Nothing
Set RsRuleObj = Nothing
%>

<script language="javaScript">

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

</script>
