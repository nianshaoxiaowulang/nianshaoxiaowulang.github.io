<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
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
<title>���ɹ鵵�����б�</title>
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
          <td width="35" align="center" alt="ˢ������" onClick="CompareDate();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
          <td>&nbsp; </td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td height="80"> <div align="center">����ʱ��ѡ��&nbsp;&nbsp; 
          <input name="FromDate" readonly type="text" size="20">
        &nbsp;&nbsp;&nbsp;&nbsp; 
        <input name="sdaf" type="button" id="sdaf" onClick="OpenWindowAndSetValue('../../FunPages/SelectDate.asp',280,120,window,document.DateForm.FromDate);" value="ѡ������">
        ----
          <input name="TentDate" readonly type="text" size="20">
        &nbsp;&nbsp;&nbsp;&nbsp; 
        <input name="sdaf" type="button" id="sdaf" onClick="OpenWindowAndSetValue('../../FunPages/SelectDate.asp',280,120,window,document.DateForm.TentDate);" value="ѡ������">
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
	if (FromDateTime=='Error') {alert('��ʼʱ�����Ͳ���ȷ');return;}
	if (TentDateTime!='')
	{
		TentDateTime=stringToDate(TentDateTime);
		if (TentDateTime=='Error') {alert('����ʱ�����Ͳ���ȷ');return;}
		if (FromDateTime>TentDateTime) {alert('��ʼʱ�䲻�����ڽ���ʱ��!');return;}
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