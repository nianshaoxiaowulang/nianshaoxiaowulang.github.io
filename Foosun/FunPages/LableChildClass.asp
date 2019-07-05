<% Option Explicit %>
<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Const.asp" -->
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
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>选择子栏目标签属性</title>
<link rel="stylesheet" href="../../CSS/ModeWindow.css">
<script language="JavaScript" src="../SysJS/PublicJS.js"></script>
</head>
<body topmargin="0" leftmargin="0" scroll=no>
<div align="center"> 
  <table width="96%" border="0" cellpadding="0" cellspacing="0">
    <tr> 
      <td width="50%" height="30">栏目数量 
        <input name="ClassNumber" id="ClassNumber" onBlur="CheckNumber(this,'栏目数量');" type="text"  style="width:70%;" value="10"> 
      </td>
      <td height="30"><div align="left">新闻数量 
          <input name="NewsNumber" id="NewsNumber" onBlur="CheckNumber(this,'新闻数量');" type="text"  style="width:70%;" value="10">
        </div></td>
    </tr>
    <tr> 
      <td height="30" colspan="2">新闻分隔 
        <input type="text" readonly  style="width:63%;" id="CompatPic" name="CompatPic"> 
        <input name="Submitff" onClick="OpenWindowAndSetValue('SelectPic.asp?CurrPath=/<% = UpFiles %>&ShowVirtualPath=true',550,290,window,document.all.CompatPic);" type="button" id="Submitff" value="选择图片"> 
        <div align="left"></div></td>
    </tr>
    <tr> 
      <td height="30" colspan="2">栏目分隔 
        <input type="text" readonly  style="width:63%;" id="ClassBGPic2" name="ClassBGPic"> 
        <input type="button" name="Submit3" onClick="OpenWindowAndSetValue('SelectPic.asp?CurrPath=/<% = UpFiles %>&ShowVirtualPath=true',550,290,window,document.all.ClassBGPic);" value="选择图片"></td>
    </tr>
    <tr> 
      <td height="30" colspan="2">导航图片 
        <input type="text" readonly  style="width:63%;" id="NaviPic2" name="NaviPic"> 
        <input type="button" name="Submit4" onClick="OpenWindowAndSetValue('SelectPic.asp?CurrPath=/<% = UpFiles %>&ShowVirtualPath=true',550,290,window,document.all.NaviPic);" value="选择图片"> 
      </td>
    <tr> 
      <td height="30" colspan="2">文字导航 
        <input type="text" name="TxtNavi" style="width:85%;"></td>
    </tr>
    <tr> 
      <td height="30">栏目间距 
        <input name="ClassRowHeight" onBlur="CheckNumber(this,'栏目间距');" type="text" id="ClassRowHeight"  style="width:70%;" value="20"></td>
      <td height="30"><div align="left">新闻行距 
          <input type="text"  style="width:70%;" onBlur="CheckNumber(this,'新闻行距');" value="20" name="NewsRowHeight" id="NewsRowHeight">
        </div></td>
    </tr>
    <tr> 
      <td height="30">栏目列数 
        <input name="ClassRowNumber" type="text" id="ClassRowNumber"  style="width:70%;" onBlur="CheckNumber(this,'栏目列数');" value="1"></td>
      <td height="30"><div align="left">新闻列数 
          <input type="text" onBlur="CheckNumber(this,'新闻列数');"  style="width:70%;" value="1" name="NewsRowNumber" id="NewsRowNumber">
        </div></td>
    </tr>
    <tr> 
      <td height="30">日期格式 
        <select style="width:70%;" name="DateRule" id="DateRule">
          <option selected>选择日期格式</option>
          <option value="1">2003-9-1</option>
          <option value="2">2003.9.1</option>
          <option value="3">2003/9/1</option>
          <option value="4">9/1/2003</option>
          <option value="5">1/9/2004</option>
          <option value="6">9-1-2004</option>
          <option value="7">9.1.2004</option>
          <option value="8">9-1</option>
          <option value="9">9/1</option>
          <option value="10">9.1</option>
          <option value="11">9月1</option>
          <option value="12">1日11时</option>
          <option value="13">1日11点</option>
          <option value="14">11时11分</option>
          <option value="15">11:11</option>
          <option value="16">2004年9月1日</option>
        </select> </td>
      <td height="30"><div align="left">日期对齐 
          <select  style="width:70%;" name="DateRight">
            <option value="Right">右对齐</option>
            <option value="Left" selected>左对齐</option>
            <option value="Center">居中</option>
          </select>
        </div></td>
    </tr>
    <tr> 
      <td height="30">更多链接 
        <select name="MoreLinkType" style="width:70%;">
          <option value="1">图片</option>
          <option value="0" selected>文字</option>
        </select></td>
      <td height="30">链接内容 
        <input title="图片地址" type="text"  style="width:70%;" name="MoreLinkContent"></td>
    </tr>
    <tr> 
      <td height="30">标题字数 
        <input name="TitleNumber" onBlur="CheckNumber(this,'标题字数');" id="TitleNumber" type="text"  style="width:70%;" value="30"> 
      </td>
      <td height="30">弹出窗口 
        <select  style="width:70%;" name="OpenType">
          <option value="0" selected>否</option>
          <option value="1">是</option>
        </select></td>
    </tr></tr>
    <tr> 
      <td height="30">标题样式 
        <input type="text" style="width:70%;" name="CSSStyle"></td>
      <td height="30">日期样式 
        <input type="text" style="width:70%;" name="DateCSSStyle"> </td>
    </tr>
    <tr> 
      <td height="30" colspan="2"> <div align="right"> 
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td>&nbsp;</td>
              <td width="100"> <div align="center"> 
                  <input name="Submitsss" type="button" id="Submitsss4" onClick="InsertScript();" value=" 确 定 ">
                </div></td>
              <td width="100"> <div align="center"> 
                  <input type="button" onClick="window.close();" name="Submit2" value=" 取 消 ">
                </div></td>
              <td>&nbsp;</td>
            </tr>
          </table>
        </div></td>
    </tr>
  </table>
</div>
</body>
</html>
<%
Set Conn = Nothing
%>
<script>
function InsertScript()
{
	var TempStr='';
	var ClassNumberStr='';
	if (document.all.ClassNumber.value=='') ClassNumberStr='10';
	else ClassNumberStr=document.all.ClassNumber.value;
	var NewsNumberStr='';
	if (document.all.NewsNumber.value=='') NewsNumberStr='10';
	else NewsNumberStr=document.all.NewsNumber.value;
	var CompatPicStr=document.all.CompatPic.value;
	var NaviPicStr=document.all.NaviPic.value;
	var ClassRowHeightStr='';
	if (document.all.ClassRowHeight.value=='') ClassRowHeightStr='20';
	else ClassRowHeightStr=document.all.ClassRowHeight.value;
	var NewsRowHeightStr='';
	if (document.all.NewsRowHeight.value=='') NewsRowHeightStr='20';
	else NewsRowHeightStr=document.all.NewsRowHeight.value;
	var ClassRowNumberStr='';
	if (document.all.ClassRowNumber.value=='') ClassRowNumberStr='1';
	else ClassRowNumberStr=document.all.ClassRowNumber.value;
	var NewsRowNumberStr='';
	if (document.all.NewsRowNumber.value=='') NewsRowNumberStr='1';
	else NewsRowNumberStr=document.all.NewsRowNumber.value;
	var DateRuleStr=document.all.DateRule.value;
	var DateRightStr='';
	DateRightStr=document.all.DateRight.value;
	var TitleNumberStr='';
	if (document.all.TitleNumber.value=='') TitleNumberStr='10';
	else TitleNumberStr=document.all.TitleNumber.value;
	var MoreLinkTypeStr=document.all.MoreLinkType.value;
	var MoreLinkContentStr=document.all.MoreLinkContent.value;
	var ClassBGPicStr=document.all.ClassBGPic.value;
	var CSSStyleStr=document.all.CSSStyle.value;
	var OpenTypeStr=document.all.OpenType.value;
	var DateCSSStyleStr=document.all.DateCSSStyle.value;
	var TxtNaviStr=document.all.TxtNavi.value;
	TempStr='{%=ChildClassList("'+ClassNumberStr+'","'+NewsNumberStr+'","'+CompatPicStr+'","'+NaviPicStr+'","'+ClassRowHeightStr+'","'+NewsRowHeightStr+'","'+ClassRowNumberStr+'","'+NewsRowNumberStr+'","'+DateRuleStr+'","'+DateRightStr+'","'+TitleNumberStr+'","'+MoreLinkTypeStr+'","'+MoreLinkContentStr+'","'+ClassBGPicStr+'","'+CSSStyleStr+'","'+OpenTypeStr+'","'+DateCSSStyleStr+'","'+TxtNaviStr+'")%}';
	window.returnValue=TempStr;
	window.close();
}
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
</script>
