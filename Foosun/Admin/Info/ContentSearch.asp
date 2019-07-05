<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
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
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P010506") then Call ReturnError()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<script language="JavaScript">
var ParentLocationStr=dialogArguments.location.href;
function GetSearchKeyWord(LocationStr,SearchStr)
{
	var SearchLocation=LocationStr.lastIndexOf(SearchStr);
	if (SearchLocation!=-1)
	{
		var StartLoc=LocationStr.indexOf('=',SearchLocation);
		var EndLoc=LocationStr.indexOf('&',SearchLocation);
		if (StartLoc!=-1)
		{
			if (EndLoc!=-1)	return LocationStr.slice(StartLoc+1,EndLoc);
			else return LocationStr.slice(StartLoc+1);
		}
		else return '';
	}
	else return '';
}
var DocumentReadyTF=false;
function document.onreadystatechange()
{
	if (DocumentReadyTF) return;
	SetValue();
	DocumentReadyTF=true;
}
function SetValue()
{
	var i=0;
	document.FunctionForm.SearchContent.value=GetSearchKeyWord(ParentLocationStr,'SearchContent');
	document.FunctionForm.SearchBeginTime.value=GetSearchKeyWord(ParentLocationStr,'SearchBeginTime');
	document.FunctionForm.SearchEndTime.value=GetSearchKeyWord(ParentLocationStr,'SearchEndTime');
	for (i=0;i<document.FunctionForm.SearchScope.options.length;i++) if (document.FunctionForm.SearchScope.options(i).value==GetSearchKeyWord(ParentLocationStr,'SearchScope')) document.FunctionForm.SearchScope.options(i).selected=true;
	for (i=0;i<document.FunctionForm.SearchType.options.length;i++) if (document.FunctionForm.SearchType.options(i).value==GetSearchKeyWord(ParentLocationStr,'SearchType')) document.FunctionForm.SearchType.options(i).selected=true;
}
</script>
<body topmargin="0" leftmargin="0">
<div align="center">
  <table width="90%" border="0" cellpadding="0" cellspacing="0">
    <form action="" method="post" name="FunctionForm">
      <tr> 
        <td width="80">搜索目标 </td>
        <td height="30"><select style="width:100%;" name="SearchScope">
            <option value="All">全部</option>
            <option value="News">新闻</option>
            <option value="DownLoad">下载</option>
          </select></td>
      </tr>
      <tr> 
        <td>搜索类型</td>
        <td height="30"><select style="width:100%;" name="SearchType">
            <option value="Title">标题</option>
            <option value="Content">内容</option>
            <option value="KeyWords">关键字</option>
          </select></td>
      </tr>
      <tr> 
        <td>搜索内容</td>
        <td height="30"><input style="width:100%;" type="text" value="" name="SearchContent"></td>
      </tr>
      <tr> 
        <td>开始日期</td>
        <td height="30"><input style="width:100%;" onFocus="OpenWindowAndSetValue('../../FunPages/SelectDate.asp',280,120,window,this);" name="SearchBeginTime" value="" readonly type="text" size="19" maxlength="20">
        </td>
      </tr>
      <tr> 
        <td>结束时间</td>
        <td height="30"> 
          <input style="width:100%;" onFocus="OpenWindowAndSetValue('../../FunPages/SelectDate.asp',280,120,window,this);" name="SearchEndTime" value="" readonly type="text" size="19" maxlength="20">
        </td>
      </tr>
      <tr> 
        <td height="30" colspan="2"><div align="center"> 
            <input name="Submit" type="button" class="SearchBtnStyle" onClick="dialogArguments.SearchSubmit(document.FunctionForm);window.close();" value=" 确 定 ">
          </div></td>
      </tr>
    </form>
  </table>
</div>
</body>
</html>
<%
Set Conn = Nothing
%>