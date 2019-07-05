<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<%
	Dim DBC,Conn
	On Error Resume Next
	Set DBC = New DataBaseClass
	Set Conn = DBC.OpenConnection()
	Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P030900") then Call ReturnError1()
Dim LableSql,RsLableObj,LableID
LableSql = "Select * from FS_LableBackUp "
Set RsLableObj = Server.CreateObject(G_FS_RS)
RSLableObj.open LableSql,conn,1,1
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>标签列表</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<script language="JavaScript" src="../../SysJS/ContentMenu.js"></script>
<body topmargin="2" leftmargin="2" onClick="SelectLable();" onselectstart="return false;">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="查看" onClick="BrowLableBack();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">查看</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35  align="center" alt="删除" onClick="DelLableBack();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">删除</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="还原" onClick="RevertLable();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">还原</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
		  <td>&nbsp; </td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
  	<td colspan="5" height="2"></td>
  </tr>
  <tr>
  <td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
	<td width="43%" height="26" class="ButtonListLeft"> 
      <div align="left">名称</div></td>
    <td width="6%" height="20" class="ButtonList"> 
      <div align="center">类型</div></td>
    <td width="22%" height="20" class="ButtonList"> 
      <div align="center">描述</div></td>
	<td width="29%" height="20" class="ButtonList"> 
      <div align="center">备份时间</div></td>
  </tr>
<%
do while not RsLableObj.eof
%>
  <tr style="background:white;cursor:default;"> 
    <td height="25"><table border="0" cellpadding="0" cellspacing="0">
		<tr>
			    <td><img src="../../Images/Lable.gif" width="18" height="18"></td>
			<td><span LableID="<%= RsLableObj("ID")%>"><% = RsLableObj("LableName")%></span></td>
		</tr></table>
	</td>
	<td><div align="center">标签</div></td>
          <td height="20"> <div align="center">
              <% = RsLableObj("Description") %>
            </div></td>
	<td height="20"><div align="center">
              <% = RsLableObj("BackUpTime") %>
            </div></tr>
<%
	RsLableObj.MoveNext
loop
RsLableObj.Close
%>
</table>
</td>
</tr>
</table>
</body>
</html>
<%
Set RsLableObj = Nothing
Set Conn = Nothing
%>
<script language="JavaScript">
var SelectedObj=null;

var DocumentReadyTF=false;
var ListObjArray = new Array();
var ContentMenuArray=new Array();
function document.onreadystatechange()
{
	if (DocumentReadyTF) return;
	IntialListObjArray();
	InitialContentListContentMenu();
	DocumentReadyTF=true;
}
function InitialContentListContentMenu()
{
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.RevertLable();",'还原','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.DelLableBack();",'删除','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.BrowLableBack();','查看','disabled');
	//ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	//ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('top.GetEkMainObject().location.reload();','刷新','');
	//ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('prompt(\'本页面路径属性\',\'<%=Request.ServerVariables("SCRIPT_NAME")%>\');','路径属性','');
}
function ContentMenuFunction(ExeFunction,Description,EnabledStr)
{
	this.ExeFunction=ExeFunction;
	this.Description=Description;
	this.EnabledStr=EnabledStr;
}
function ContentMenuShowEvent()
{

	ChangeContentMenuStatus();		
}
function ChangeContentMenuStatus()
{
	var EventObjInArray=false,SelectContent='',DisabledContentMenuStr='';
	if (SelectedObj!=null)
	{
		if (SelectedObj.LableID!=null) DisabledContentMenuStr='';
		else
		{
			DisabledContentMenuStr=',还原,删除,查看,'
		}
	}
	else
	{
		DisabledContentMenuStr=',还原,删除,查看,'
	}
	for (var i=0;i<ContentMenuArray.length;i++)
	{
		if (DisabledContentMenuStr.indexOf(ContentMenuArray[i].Description)!=-1) ContentMenuArray[i].EnabledStr='disabled';
		else  ContentMenuArray[i].EnabledStr='';
	}
}
function FolderFileObj(Obj,Index,Selected)
{
	this.Obj=Obj;
	this.Index=Index;
	this.Selected=Selected;
}
function IntialListObjArray()
{
	var CurrObj=null,j=1;
	for (var i=0;i<document.all.length;i++)
	{
		CurrObj=document.all(i);
		if (CurrObj.StyleID!=null)
		{
			ListObjArray[ListObjArray.length]=new FolderFileObj(CurrObj,j,false);
			j++;
		}
	}
}
function SelectLable()
{
	var el=event.srcElement;
	if (SelectedObj!=null) {SelectedObj.className='TempletItem';SelectedObj=null;} 
	if (el.LableID!=null) {SelectedObj=el;el.className='TempletSelectItem';}
}
function BrowLableBack()
{
	if (SelectedObj==null) {alert('请选择标签');return;}
	if (SelectedObj.LableID!=null) {OpenWindow('Frame.asp?PageTitle=查看标签&FileName=LableContent.asp&ID='+SelectedObj.LableID,360,220,window);return;}
}
function DelLableBack()
{
	if (SelectedObj==null) {alert('请选择标签');return;}
	if (SelectedObj.LableID!=null) {OpenWindow('Frame.asp?PageTitle=删除备份标签&FileName=DelLableBackUp.asp&ID='+SelectedObj.LableID,190,110,window);return;}
}
function RevertLable()
{
	if (SelectedObj==null) {alert('请选择标签');return;}
	if (SelectedObj.LableID!=null) {OpenWindow('Frame.asp?PageTitle=还原标签&FileName=RevertLable.asp&LableID='+SelectedObj.LableID,260,110,window);return;}
}
</script>