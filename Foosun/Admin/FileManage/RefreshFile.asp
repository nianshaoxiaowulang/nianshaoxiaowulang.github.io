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
Dim DBC,Conn,RecordConn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
DBC.ConnStr = "DBQ=" + server.mappath(RecordDataBaseConnectStr) + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
Set RecordConn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P070607") then Call ReturnError1()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>生成归档新闻列表</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<body topmargin="0" leftmargin="0" scroll=no>
<form action="RefreshFileSave.asp" method="post" name="DateForm">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="35" align="center" alt="刷新新闻" onClick="CompareDate();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">生成</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
          <td>&nbsp; </td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td height="80"> <div align="center">生成时间选择&nbsp;&nbsp; 
          <input name="FromDate" readonly type="text" size="20">
        &nbsp;&nbsp;&nbsp;&nbsp; 
        <input name="sdaf" type="button" id="sdaf" onClick="OpenWindowAndSetValue('../../FunPages/SelectDate.asp',280,120,window,document.DateForm.FromDate);" value="选择日期">
        ----
          <input name="TentDate" readonly type="text" size="20">
        &nbsp;&nbsp;&nbsp;&nbsp; 
        <input name="sdaf" type="button" id="sdaf" onClick="OpenWindowAndSetValue('../../FunPages/SelectDate.asp',280,120,window,document.DateForm.TentDate);" value="选择日期">
      </div></td>
  </tr>
</table>
</form>
</body>
</html>
<%
Set Conn = Nothing
Set RecordConn = Nothing
%>
<script language="JavaScript">
function CompareDate()
{   
	var FromDateTime = document.DateForm.FromDate.value;
	var TentDateTime = document.DateForm.TentDate.value;
	FromDateTime=stringToDate(FromDateTime);
	if (FromDateTime=='Error') {alert('开始时间类型不正确');return;}
	if (TentDateTime!='')
	{
		TentDateTime=stringToDate(TentDateTime);
		if (TentDateTime=='Error') {alert('结束时间类型不正确');return;}
		if (FromDateTime>TentDateTime) {alert('开始时间不能晚于结束时间!');return;}
	}
	document.DateForm.submit();
} 

function isDateString(sDate)
{	var iaMonthDays = [31,28,31,30,31,30,31,31,30,31,30,31]
	var iaDate = new Array(3)
	var year, month, day
	if (arguments.length != 1) return false
	iaDate = sDate.toString().split("-")
	if (iaDate.length != 3) return false
	if (iaDate[1].length > 2 || iaDate[2].length > 2) return false
	if (isNaN(iaDate[0])||isNaN(iaDate[1])||isNaN(iaDate[2])) return false
	year = parseFloat(iaDate[0])
	month = parseFloat(iaDate[1])
	day=parseFloat(iaDate[2])
	if (year < 1900 || year > 2100) return false
	if (((year % 4 == 0) && (year % 100 != 0)) || (year % 400 == 0)) iaMonthDays[1]=29;
	if (month < 1 || month > 12) return false
	if (day < 1 || day > iaMonthDays[month - 1]) return false
	return true
}

function stringToDate(sDate)
{	var bValidDate, year, month, day
	var iaDate = new Array(3)
	bValidDate = isDateString(sDate)
	if (bValidDate)
	{  iaDate = sDate.toString().split("-")
		year = parseFloat(iaDate[0])
		month = parseFloat(iaDate[1]) - 1
		day=parseFloat(iaDate[2])
		return (new Date(year,month,day))
	}
	else return 'Error';
} 
</script>