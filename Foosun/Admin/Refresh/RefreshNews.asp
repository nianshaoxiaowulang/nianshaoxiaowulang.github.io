<% Option Explicit %>
<!--#include file="Function.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P030400") then Call ReturnError1()
Session("AllClassID")=""
Session("NewsTotalNum")=""
Dim TempClassListStr
	TempClassListStr = ClassList("ClassID")

%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�������ɹ���</title>
</head>
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<body topmargin="2" leftmargin="2" oncontextmenu="return false;">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="28" class="ButtonListLeft">
<div align="center"><strong>����ҳ���ɹ���</strong></div></td>
</tr>
</table>
<table width="100%"  border="0" cellspacing="8" cellpadding="0">
  <tr> 
    <td width="10%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
    <td width="78%" colspan="2">&nbsp;</td>
  </tr>
  <form action="RefreshNewsSave.asp?Types=DatesType" method="post" name="DateForm">
    <tr> 
      <td>&nbsp;</td>
      <td>����������</td>
      <td colspan="2"><input name="FromDate" type="text" id="FromDate" readonly style="width:20%" value="<%=Date()%>"> 
        <input type="button" name="Submit4" value="ѡ������" onClick="OpenWindowAndSetValue('../../FunPages/SelectDate.asp',280,110,window,document.DateForm.FromDate);">
        �� 
        <input name="TentDate" type="text" id="TentDate" readonly style="width:20%" value="<%=Date()%>"> 
        <input type="button" name="Submit4" value="ѡ������" onClick="OpenWindowAndSetValue('../../FunPages/SelectDate.asp',280,110,window,document.DateForm.TentDate);"> 
        <input name="imageField4" type="image" src="../../Images/Publish.gif" width="75" height="21" border="0" onClick="CompareDate();"> 
      </td>
    </tr>
  </form>
  <form action="RefreshNewsSave.asp?Types=NewType" method="post" name="NewForm">
    <tr> 
      <td>&nbsp;</td>
      <td> �������� </td>
      <td><input name="NewNum"  type="text" id="NewNum" style="width:20%" value="10"> 
        <input name="imageField2" type="image" src="../../Images/Publish.gif" width="75" height="21" border="0" onClick="SubmitLastType();"> 
      </td>
      <td>&nbsp;</td>
    </tr>
  </form>
  <form action="RefreshNewsSave.asp?Types=AllType" method="post" name="AllForm">
    <tr> 
      <td>&nbsp;</td>
      <td> ����������Ϣ </td>
      <td colspan="2"><input name="imageField" type="image" src="../../Images/Publish.gif" width="75" height="21" border="0"> 
      </td>
    </tr>
  </form>
  <form action="RefreshNewsSave.asp?Types=ClassType" method="post" name="ClassForm">
    <tr> 
      <td>&nbsp;</td>
      <td> ����Ŀ���� </td>
      <td colspan="2"><select name="ClassID" size="10" multiple style="width:80%">
          <% =TempClassListStr %>
        </select>
        <br>
        <input name="imageField3" type="image" onClick="AccordClassRefresh();" src="../../Images/Publish.gif" width="75" height="21" border="0"> 
        <input name="IssueSubClass" type="checkbox" id="IssueSubClass3" value="IssueSubClass">
        ��������Ŀ </td>
    </tr>
  </form>
  <tr> 
    <td>&nbsp;</td>
    <td colspan="3"><font color=red>ע��:�����ɹ����У�����ˢ�´�ҳ��</font></td>
  </tr>
</table>
</body>
</html>
<script>
function CompareDate()
{   
	var FromDateTime = document.DateForm.FromDate.value;
	var TentDateTime = document.DateForm.TentDate.value;
	FromDateTime=stringToDate(FromDateTime);
	if (FromDateTime=='Error') {alert('��ʼʱ�����Ͳ���ȷ');return;}
	TentDateTime=stringToDate(TentDateTime);
	if (TentDateTime=='Error') {alert('����ʱ�����Ͳ���ȷ');return;}
	if (FromDateTime>TentDateTime) alert('��ʼʱ�䲻�����ڽ���ʱ��!');
	else document.DateForm.submit();
} 

function SubmitLastType()
{
	if (document.NewForm.NewNum.value=='') {alert('����д��������');document.NewForm.NewNum.focus();}
	else document.NewForm.submit();
}
 
function AccordClassRefresh()
{
	//if (document.ClassForm.NumClass.value=='') {alert('����д��������');document.ClassForm.NumClass.focus();}
	 document.ClassForm.submit();
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