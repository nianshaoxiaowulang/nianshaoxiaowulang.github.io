<% Option Explicit %>
<!--#include file="../../Inc/Const.asp" -->
<!--#include file="../../Inc/Cls_DB.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../Inc/Session.asp" -->
<!--#include file="../../Inc/CheckPopedom.asp" -->
<!--#include file="../../Inc/Function.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P030800") then Call ReturnError()
Dim TempClassListStr
	TempClassListStr = ClassList("ClassEName")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../../CSS/ModeWindow.css">
<script language="JavaScript" src="../SysJS/PublicJS.js"></script>
<title>焦点图片</title>
</head>
<body topmargin="0" leftmargin="0">
<table width="96%" border="0" align="center" cellpadding="3" cellspacing="0">
  <tr> 
    <td width="456"><div align="left">栏目列表 
        <select name="ClassList" id="ClassList" style="width:70%;">
          <option value="" selected>栏目选择</option>
          <% = TempClassListStr %>
        </select>
      </div></td>
    <td width="425"><div align="left">新闻数量 
        <input name="NewsNumber" type="text" id="NewsNumber" style="width:70%;" value="10">
      </div></td>
  </tr>
  <tr> 
    <td><div align="left">调用子类 
        <select name="HaveChild" id="HaveChild" style="width:70%;">
          <option value="1" selected>是</option>
          <option value="0">否</option>
        </select>
      </div></td>
    <td><div align="left">排列列数 
        <input name="RowNumber" type="text" id="RowNumber" style="width:70%;" value="1">
      </div></td>
  </tr>
  <tr> 
    <td><div align="left">显示标题 
        <select name="IsTitle" id="IsTitle" style="width:70%;">
          <option value="1">是</option>
          <option value="0" selected>否</option>
        </select>
      </div></td>
    <td><div align="left">标题字数 
        <input name="TitleNumber" type="text" id="TitleNumber" style="width:70%; " value="30">
      </div></td>
  </tr>
  <tr> 
    <td><div align="left">标题样式 
        <input type="text" style="width:70%;" name="CssFile" id="CssFile">
      </div></td>
    <td><div align="left">弹出窗口 
        <select name="OpenMode" id="OpenMode" style="width:70%;">
          <option value="1" selected>是</option>
          <option value="0">否</option>
        </select>
      </div></td>
  </tr>
  <tr> 
    <td><div align="left">图片规格 
        <input name="PicWidth" type="text" id="SizeHeight2" style="width:30%;" value="60">
        * 
        <input name="PicHeight" type="text" id="SizeWidth2" style="width:30%;" value="60">
      </div></td>
    <td> <div align="left">间　　距 
        <input name="RowSpace" type="text" id="RowSpace" style="width:70%; " value="6">
      </div></td>
  </tr>
  <tr> 
    <td height="30" colspan="2"><div align="center"> 
        <input type="button" onClick="InsertScript();" name="Submit" value=" 确 定 ">
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="button" onClick="window.close();" name="Submit2" value=" 取 消 ">
      </div></td>
  </tr>
</table>
</body>
</html>
<script language="JavaScript">
function InsertScript()
{
	var NewsNumberStr='';
	if (document.all.NewsNumber.value=='') NewsNumberStr='10';
	else NewsNumberStr=document.all.NewsNumber.value;
	var ClassListStr=document.all.ClassList.value;
	var HaveChildStr=document.all.HaveChild.value;
	var IsTitleStr=document.all.IsTitle.value;
	var TitleNumberStr='';
	if (document.all.TitleNumber.value=='') TitleNumberStr='20';
	else TitleNumberStr=document.all.TitleNumber.value;
	var OpenModeStr=document.all.OpenMode.value;
	var RowNumberStr='';
	if (document.all.RowNumber.value=='') RowNumberStr='1';
	else RowNumberStr=document.all.RowNumber.value;
	var PicWidthStr='';
	if (document.all.PicWidth.value=='') PicWidthStr='60';
	else PicWidthStr=document.all.PicWidth.value;
	var PicHeightStr='';
	if (document.all.PicHeight.value=='') PicHeightStr='60';
	else PicHeightStr=document.all.PicHeight.value;
	var CssFileStr=document.all.CssFile.value;
	var RowSpaceStr='';
	if (document.all.RowSpace.value=='') RowSpaceStr='20';
	else RowSpaceStr=document.all.RowSpace.value;
	window.returnValue='{%=FocusPic("'+ClassListStr+'","'+NewsNumberStr+'","'+HaveChildStr+'","'+IsTitleStr+'","'+TitleNumberStr+'","'+OpenModeStr+'","'+RowNumberStr+'","'+PicWidthStr+'","'+PicHeightStr+'","'+CssFileStr+'","'+RowSpaceStr+'")%}';
	window.close();
}
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
</script>