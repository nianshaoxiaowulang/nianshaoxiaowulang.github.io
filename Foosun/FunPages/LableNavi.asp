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
<title>导航属性选择</title>
</head>
<link rel="stylesheet" href="../../CSS/ModeWindow.css">
<script language="JavaScript" src="../SysJS/PublicJS.js"></script>
<body topmargin="0" leftmargin="0">
<div align="center">
  <table width="96%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="50%" height="30" nowrap> <div align="left">导航类型 
          <select name="NaviType" id="NaviType"  style="width:70%;">
            <option value="1" selected>栏目导航</option>
            <option value="2">专题导航</option>
            <option value="3">插件导航</option>
            <option value="4">栏目+专题导航</option>
            <option value="5">栏目+插件导航</option>
            <option value="6">专题+插件导航</option>
            <option value="7">栏目+专题+插件导航</option>
          </select>
        </div></td>
      <td width="50%" nowrap>排列列数 
        <input name="RowNumber" onBlur="CheckNumber(this,'显示列数');" type="text"  style="width:70%;" value="10" size="6"> 
      </td>
    </tr>
    <tr> 
      <td height="30" colspan="2"><div align="left">导航图片 
          <input type="text" readonly  style="width:63%;" id="NaviPic" name="NaviPic">
          <input type="button" name="Submit3" onClick="OpenWindowAndSetValue('SelectPic.asp?CurrPath=/<% = UpFiles %>&ShowVirtualPath=true',550,290,window,document.all.NaviPic);" value="选择图片">
        </div></td>
    </tr>
    <tr> 
      <td height="30" colspan="2">分隔图片 
        <input type="text"  style="width:63%;" readonly name="BGPic"> <input type="button" name="Submit4" onClick="OpenWindowAndSetValue('SelectPic.asp?CurrPath=/<% = UpFiles %>&ShowVirtualPath=true',550,290,window,document.all.BGPic);" value="选择图片"></td>
    </tr>
    <tr> 
      <td height="30" colspan="2">文字导航 
        <input type="text" name="TxtNavi" style="width:85%;"></td>
    </tr>
    <tr> 
      <td height="30"><div align="left">弹出窗口 
          <select style="width:70%;" name="OpenType">
            <option value="0" selected>否</option>
            <option value="1">是</option>
          </select>
        </div></td>
      <td>标题样式
<input type="text" style="width:70%;" name="CSSStyle"></td>
    </tr>
    <tr> 
      <td height="30"><div align="center"> 
          <input type="button" name="Submit" onClick="InsertScript();" value=" 确 定 ">
        </div></td>
      <td><div align="center"> 
          <input type="button" name="Submit2" onClick="window.close();" value=" 取 消 ">
        </div></td>
    </tr>
  </table>
</div>
</body>
</html>
<script language="JavaScript">
function InsertScript()
{
	var RowNumberStr='';
	var NaviPicStr='';
	var CompatPicStr='';
	var OpenTypeStr='';
	var CSSStyleStr='';
	if (document.all.RowNumber.value=='') RowNumberStr='10';
	else RowNumberStr=document.all.RowNumber.value;
	NaviPicStr=document.all.NaviPic.value;
	CompatPicStr=document.all.BGPic.value;
	OpenTypeStr=document.all.OpenType.value;
	CSSStyleStr=document.all.CSSStyle.value;
	var TxtNaviStr=document.all.TxtNavi.value;
	window.returnValue='{%=LocationNavi("'+document.all.NaviType.value+'","'+RowNumberStr+'","'+NaviPicStr+'","'+CompatPicStr+'","'+OpenTypeStr+'","'+CSSStyleStr+'","'+TxtNaviStr+'")%}';
	window.close();
}
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
</script>