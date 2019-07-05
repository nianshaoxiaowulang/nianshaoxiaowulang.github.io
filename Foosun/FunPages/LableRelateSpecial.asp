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
<%
if Not JudgePopedomTF(Session("Name"),"P030800") then Call ReturnError()
%><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>相关新闻属性</title>
</head>
<link rel="stylesheet" href="../../CSS/ModeWindow.css">
<script language="JavaScript" src="../SysJS/PublicJS.js"></script>
<body topmargin="0" leftmargin="0" scroll=no>
<div align="center">
  <table width="96%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="50%" height="30"> <div align="left">新闻数量 
          <input name="NewsNumber" type="text" id="NewsNumber" style="width:70%;" value="10">
        </div></td>
      <td>排列列数 
        <input name="RowNumber" type="text" id="RowNumber2" style="width:70%;" value="1"></td>
    </tr>
    <tr> 
      <td height="30"><div align="left">标题字数 
          <input name="TitleNumber" type="text" id="TitleNumber" style="width:70%;" value="30">
        </div></td>
      <td>标题样式
<input type="text" style="width:70%;" name="CSSStyle"></td>
    </tr>
    <tr> 
      <td height="30" colspan="2">导航图片 
        <input type="text" readonly style="width:63%;" name="NaviPic" id="NaviPic">
        <input type="button" onClick="OpenWindowAndSetValue('SelectPic.asp?CurrPath=/<% = UpFiles %>&ShowVirtualPath=true',550,290,window,document.all.NaviPic);" name="Submit3" value="选择图片"> </td>
    </tr>
    <tr> 
      <td height="30" colspan="2"><div align="left"></div>
        分隔图片 
        <input type="text" readonly style="width:63%;" id="CompatPic" name="CompatPic">
        <input type="button" onClick="OpenWindowAndSetValue('SelectPic.asp?CurrPath=/<% = UpFiles %>&ShowVirtualPath=true',550,290,window,document.all.CompatPic);" name="Submit4" value="选择图片"></td>
    </tr>
    <tr> 
      <td height="30" colspan="2"><div align="left">文字导航 
          <input type="text" name="TxtNavi" style="width:85%;">
        </div>
        </td>
    </tr>
    <tr> 
      <td height="30"><div align="center"> 
          <input type="button" onClick="InsertScript();" name="Submit" value=" 确 定 ">
        </div></td>
      <td><div align="center"> 
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
	var NewsNumberStr='';
	var TitleNumberStr='';
	var RowNumberStr='';
	var CSSStyleStr=document.all.CSSStyle.value;
	if (document.all.NewsNumber.value=='') NewsNumberStr='10';
	else NewsNumberStr=document.all.NewsNumber.value;
	if (document.all.TitleNumber.value=='') TitleNumberStr='10';
	else TitleNumberStr=document.all.TitleNumber.value;
	if (document.all.RowNumber.value=='') RowNumberStr='10';
	else RowNumberStr=document.all.RowNumber.value;
	var NaviPicStr=document.all.NaviPic.value;
	var CompatPicStr=document.all.CompatPic.value;
	var TxtNaviStr=document.all.TxtNavi.value;
	//var CompatPicStr=document.all.CompatPic.value;
	window.returnValue='{%=RelateSpecialNews("'+NewsNumberStr+'","'+TitleNumberStr+'","'+RowNumberStr+'","'+NaviPicStr+'","'+CompatPicStr+'","'+CSSStyleStr+'","'+TxtNaviStr+'")%}';
	window.close();
}
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
</script>