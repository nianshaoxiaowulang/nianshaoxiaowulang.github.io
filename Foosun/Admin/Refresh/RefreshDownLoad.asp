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
if Not JudgePopedomTF(Session("Name"),"P030600") then Call ReturnError1()
Dim TempClassListStr
	TempClassListStr = ClassList
Function ClassList()
	Dim Rs,RsNews
	Set Rs = Conn.Execute("select ClassID,ClassCName from FS_newsclass where ParentID = '0' and IsOutClass=0 order by AddTime desc")
	do while Not Rs.Eof
	    Set RsNews = Conn.Execute("Select count(DownLoadID) from FS_DownLoad where ClassID='"&Rs("ClassID")&"'")
		ClassList = ClassList & "<option value="&Rs("ClassID")&"*"&RsNews(0)&"" & ">" & Rs("ClassCName") & chr(10) & chr(13)
		ClassList = ClassList & ChildClassList(Rs("ClassID"),"")
		Rs.MoveNext	
	loop
	Rs.Close
	Set Rs = Nothing
End Function
Function ChildClassList(ClassID,Temp)
	Dim TempRs,TempStr,RsNews
	Set TempRs = Conn.Execute("Select ClassID,ClassCName,ChildNum from FS_NewsClass where ParentID = '" & ClassID & "' order by AddTime desc ")
	TempStr = Temp & " - "
	do while Not TempRs.Eof
	    Set RsNews = Conn.Execute("Select count(DownLoadID) from FS_DownLoad where ClassID='"&TempRs("ClassID")&"'")
		if TempRs("ChildNum") = 0 then
			ChildClassList = ChildClassList & "<option value="&TempRs("ClassID")&"*"&RsNews(0)&"" & ">" & TempStr & TempRs("ClassCName") & "</option>"& chr(10) & chr(13)
		else
			ChildClassList = ChildClassList & "<option value="&TempRs("ClassID")&"*"&RsNews(0)&"" & ">" & TempStr & TempRs("ClassCName") & "</option>"& chr(10) & chr(13)
		end if
		ChildClassList = ChildClassList & ChildClassList(TempRs("ClassID"),TempStr)
		TempRs.MoveNext
	loop
	TempRs.Close
	Set TempRs = Nothing
End Function
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>下载生成管理</title>
</head>
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<body topmargin="2" leftmargin="2" oncontextmenu="return false;">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="28" class="ButtonListLeft"> 
      <div align="center"><strong>新闻页生成管理</strong></div></td>
</tr>
</table>
<table width="100%"  border="0" cellspacing="8" cellpadding="0">
  <tr>
    <td width="10%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
    <td width="78%">&nbsp;</td>
  </tr>
  <form action="RefreshDownLoadSave.asp?Types=DatesType" method="post" name="DateForm">
  <tr>
    <td>&nbsp;</td>
    <td>按日期生成</td>
    <td><input name="FromDate" type="text" id="FromDate" readonly style="width:20%" value="<%=Date()%>">
      <input type="button" name="Submit4" value="选择日期" onClick="OpenWindowAndSetValue('../../FunPages/SelectDate.asp',280,110,window,document.DateForm.FromDate);">
      到
      <input name="TentDate" type="text" id="TentDate" readonly style="width:20%" value="<%=Date()%>">
      <input type="button" name="Submit4" value="选择日期" onClick="OpenWindowAndSetValue('../../FunPages/SelectDate.asp',280,110,window,document.DateForm.TentDate);">
        <img name="imageField4" src="../../Images/Publish.gif" width="75" height="21" border="0" onClick="CompareDate();"> 
      </td>
  </tr>
  </form>
  <form action="RefreshDownLoadSave.asp?Types=NewType" method="post" name="NewForm">
  <tr>
    <td>&nbsp;</td>
    <td> 生成最新 </td>
    <td><input name="NewNum" type="text" id="NewNum" style="width:20%" value="10">
        <img name="imageField2" src="../../Images/Publish.gif" width="75" height="21" border="0" onClick="SubmitLastType();"> 
      </td>
  </tr>
  </form>
  <form action="RefreshDownLoadSave.asp?Types=AllType" method="post" name="AllForm">
  <tr>
    <td>&nbsp;</td>
      <td> 生成所有下载</td>
      <td><input name="imageField" type="image" src="../../Images/Publish.gif" width="75" height="21" border="0"> 
      </td>
  </tr>
  </form>
  <form action="RefreshDownLoadSave.asp?Types=ClassType" method="post" name="ClassForm">
  <tr>
    <td>&nbsp;</td>
    <td> 按栏目生成 </td>
    <td><select name="dDClass" style="width:20%" onChange="ChooseNewsNum(this.options[this.selectedIndex].value);">
	<% =TempClassListStr %>
    </select><input type="hidden" name="ClassID">
    <input name="NumClass" onBlur="CheckNumber(this,'新闻数量');" type="text" id="NumClass" style="width:14%">
        <img name="imageField3" src="../../Images/Publish.gif" onClick="AccordClassRefresh();" width="75" height="21" border="0"> 
      </td>
  </tr>
  </form>
  <tr>
    <td>&nbsp;</td>
    <td colspan="2"><font color=red>注意:在生成过程中，请勿刷新此页面</font></td>
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
	if (FromDateTime=='Error') {alert('开始时间类型不正确');return;}
	TentDateTime=stringToDate(TentDateTime);
	if (TentDateTime=='Error') {alert('结束时间类型不正确');return;}
	if (FromDateTime>TentDateTime) alert('开始时间不能晚于结束时间!');
	else document.DateForm.submit();
}
 
function SubmitLastType()
{
	if (document.NewForm.NewNum.value=='') {alert('请填写新闻数量');document.NewForm.NewNum.focus();}
	else document.NewForm.submit();
}

function AccordClassRefresh()
{
	if (document.ClassForm.NumClass.value=='') {alert('请填写新闻数量');document.ClassForm.NumClass.focus();}
	else document.ClassForm.submit();
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
 
function ChooseNewsNum(ClassCode)
{
	var TempArray,revisedTitle,revTransferNum;
	TempArray = ClassCode.split('*');
	revisedTitle = TempArray[0];
	revTransferNum = TempArray[1];
	document.ClassForm.ClassID.value=revisedTitle;
	document.ClassForm.NumClass.value=revTransferNum;
	return;
}
 
function ShowFirst()
{
	var TempArray,revisedTitle,revTransferNum;
	var ClassCode = document.ClassForm.dDClass.value;
	TempArray = ClassCode.split('*');
	revisedTitle = TempArray[0];
	revTransferNum = TempArray[1];
	document.ClassForm.ClassID.value=revisedTitle;
	document.ClassForm.NumClass.value=revTransferNum;
}
ShowFirst();
</script>