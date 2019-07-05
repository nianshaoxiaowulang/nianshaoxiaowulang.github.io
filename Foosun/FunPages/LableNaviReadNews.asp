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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>导读新闻属性</title>
</head>
<link rel="stylesheet" href="../../CSS/ModeWindow.css">
<script language="JavaScript" src="../SysJS/PublicJS.js"></script>
<body topmargin="0" leftmargin="0" scroll=no scroll=no>
<div align="center">
  <table width="96%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td height="30"> <div align="left">栏目列表 
          <select name="ClassList" style="width:70%;">
            <option value="" selected>栏目选择</option>
            <% =TempClassListStr %>
          </select>
        </div></td>
    <td height="30"><div align="left">调用子类 
        <select name="SoonClass" style="width:70%;">
          <option value="1" selected>是</option>
          <option value="0">否</option>
        </select>
      </div></td>
    </tr> 
	<tr> 
      <td width="50%" height="30"> <div align="left">图片宽度 
          <input name="PicWidth" type="text" style="width:70%;" value="60">
        </div></td>
      <td height="30">图片高度 
        <input name="PicHeight" type="text" style="width:70%;" value="60"></td>
    </tr>
    <tr> 
      <td height="30" colspan="2"><div align="left">导航图片 
          <input type="text" readonly  style="width:63%;" id="NaviPic" name="NaviPic">
          <input type="button" onClick="OpenWindowAndSetValue('SelectPic.asp?CurrPath=/<% = UpFiles %>&ShowVirtualPath=true',550,290,window,document.all.NaviPic);" name="Submit" value="选择图片">
        </div></td>
    </tr>
    <tr> 
      <td height="30" colspan="2"><div align="left">分隔图片 
          <input type="text" readonly  style="width:63%;" id="BGPic" name="BGPic">
          <input type="button" onClick="OpenWindowAndSetValue('SelectPic.asp?CurrPath=/<% = UpFiles %>&ShowVirtualPath=true',550,290,window,document.all.BGPic);" name="Submit3" value="选择图片">
        </div></td>
    </tr>
    <tr> 
      <td height="30" colspan="2">文字导航 
        <input type="text" name="TxtNavi" style="width:85%;"></td>
    </tr>
    <tr> 
      <td height="30"><div align="left">标题字数 
          <input name="TitleNumber" id="TitleNumber" onBlur="CheckNumber(this,'标题字数');" type="text"  style="width:70%;" value="30">
        </div></td>
      <td height="30">标题样式
<input type="text"  style="width:70%;" name="CSSStyle"></td>
    </tr>
    <tr> 
      <td height="30"><div align="left">新闻数量 
          <input name="NewsNumber" id="NewsNumber" onBlur="CheckNumber(this,'新闻数量');" type="text"  style="width:70%;" value="10">
        </div></td>
      <td height="30">排列列数 
        <input name="RowNumber" type="text" onBlur="CheckNumber(this,'新闻列数');" id="RowNumber"  style="width:70%;" value="1"></td>
    </tr>
    <tr> 
      <td height="30"><div align="left">弹出窗口 
          <select  style="width:70%;" name="OpenType">
            <option value="0" selected>否</option>
            <option value="1">是</option>
          </select>
        </div></td>
      <td height="30"> 新闻行距 
        <input name="RowHeight" type="text" style="width:70%;" id="RowHeight" value="20"></td>
    </tr>
    <tr> 
      <td height="30"><div align="center"> 
          <input name="Submitafd" type="button" onClick="InsertScript();" id="Submitafd" value=" 确 定 ">
        </div></td>
      <td height="30"><div align="center"> 
          <input type="button" onClick="window.close();" name="Submit2" value=" 取 消 ">
        </div></td>
    </tr>
  </table>
</div>
</body>
</html>
<script language="JavaScript">
function InsertScript()
{
	var PicWidthStr='';
	if (document.all.PicWidth.value=='') PicWidthStr='60';
	else PicWidthStr=document.all.PicWidth.value;
	var PicHeightStr='';
	if (document.all.PicHeight.value=='') PicHeightStr='60';
	else PicHeightStr=document.all.PicHeight.value;
	var NaviPicStr=document.all.NaviPic.value;
	var BGPicStr=document.all.BGPic.value;
	var ClassListStr=document.all.ClassList.value;
	var SoonClassStr=document.all.SoonClass.value;
	var TitleNumberStr='';
	if (document.all.TitleNumber.value=='') TitleNumberStr='10';
	else TitleNumberStr=document.all.TitleNumber.value;
	var CSSStyleStr=document.all.CSSStyle.value;
	var NewsNumberStr='';
	if (document.all.NewsNumber.value=='') NewsNumberStr='10';
	else NewsNumberStr=document.all.NewsNumber.value;
	var RowNumberStr='';
	if (document.all.RowNumber.value=='') RowNumberStr='10';
	else RowNumberStr=document.all.RowNumber.value;
	var OpenTypeStr=document.all.OpenType.value;
	var RowHeightStr='';
	if (document.all.RowHeight.value=='') RowHeightStr='20';
	else RowHeightStr=document.all.RowHeight.value;
	var TxtNaviStr=document.all.TxtNavi.value;
	window.returnValue='{%=NaviReadNews("'+ClassListStr+'","'+SoonClassStr+'","'+PicWidthStr+'","'+PicHeightStr+'","'+NaviPicStr+'","'+BGPicStr+'","'+TitleNumberStr+'","'+CSSStyleStr+'","'+NewsNumberStr+'","'+RowNumberStr+'","'+OpenTypeStr+'","'+RowHeightStr+'","'+TxtNaviStr+'")%}';
	window.close();
}
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
</script>