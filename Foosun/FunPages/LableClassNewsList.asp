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
<%
if Not JudgePopedomTF(Session("Name"),"P030800") then Call ReturnError()
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>栏目新闻列表属性</title>
</head>
<link rel="stylesheet" href="../../CSS/ModeWindow.css">
<script language="JavaScript" src="../SysJS/PublicJS.js"></script>
<body topmargin="0" leftmargin="0">
<div align="center">
  <table width="90%" border="0" cellspacing="0" cellpadding="0">
    <form action="" method="post" name="CNListForm">
      <tr> 
        <td width="42%" height="30">分页数量 
          <input name="NewsNumber" type="text" id="NewsNumber2" style="width:70%;" value="10"> 
        </td>
        <td width="50%"><div align="left">新闻行距 
            <input name="RowHeight" type="text" style="width:70%;" value="20">
          </div></td>
      </tr>
      <tr> 
        <td height="30"><div align="left">排列列数 
            <input name="RowNumber" type="text" id="RowNumber" style="width:70%;" value="1">
          </div></td>
        <td height="30"><div align="left">标题字数 
            <input name="TitleNumber" type="text" id="TitleNumber" style="width:70%;" value="40">
          </div></td>
      </tr>
      <tr> 
        <td height="30" colspan="2"><div align="left">导航图片
            <input name="NaviPic" readonly type="text" id="NaviPic" style="width:63%;">
            <input type="button" name="Subfdgf" value="选择图片" onClick="OpenWindowAndSetValue('SelectPic.asp?CurrPath=/<% = UpFiles %>&ShowVirtualPath=true',550,290,window,document.CNListForm.NaviPic);">
          </div></td>
      </tr>
      <tr> 
        <td height="30" colspan="2"><div align="left">分隔图片 
            <input name="BGPic" readonly type="text" id="BGPic" style="width:63%;">
            <input type="button" name="sdafsdf" value="选择图片" onClick="OpenWindowAndSetValue('SelectPic.asp?CurrPath=/<% = UpFiles %>&ShowVirtualPath=true',550,290,window,document.CNListForm.BGPic);">
          </div></td>
      </tr>
      <tr> 
        <td height="30" colspan="2">文字导航 
          <input type="text" name="TxtNavi" style="width:85%;"></td>
      </tr>
      <tr> 
        <td height="30"><div align="left">标题样式 
            <input type="text" style="width:70%;" name="CssFile" id="CssFile">
          </div></td>
        <td height="30"><div align="left">日期样式 
            <input type="text" style="width:70%;" name="DateCSSStyle">
          </div></td>
      </tr>
      <tr> 
        <td height="30"><div align="left">日期格式 
            <select  style="width:70%;" name="DateRule" id="DateRule">
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
            </select>
          </div></td>
        <td height="30"><div align="left">日期对齐 
            <select  style="width:70%;" name="DateRight">
              <option value="Right">右对齐</option>
              <option value="Left" selected>左对齐</option>
              <option value="Center">居中</option>
            </select>
          </div></td>
      </tr>
      <tr> 
        <td height="30"><div align="left">弹出窗口 
            <select name="OpenMode" id="OpenMode" style="width:70%">
              <option value="1">是</option>
              <option value="0" selected>否</option>
            </select>
          </div></td>
        <td height="30"><div align="left">新闻分页 
            <select name="DetachPage" id="DetachPage" style="width:70%">
              <option value="1" selected>是</option>
              <option value="0">否</option>
            </select>
          </div></td>
      </tr>
      <tr> 
        <td height="30" colspan="2"><div align="center"> 
            <input type="button" onClick="InsertScript();" name="Submit" value=" 确 定 ">
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
            <input type="button" onClick="window.close();" name="Submit2" value=" 取 消 ">
          </div></td>
      </tr>
    </form>
  </table>
</div>
</body>
</html>
<%
Set Conn = Nothing
%>
<script language="JavaScript">
function InsertScript()
{
	var ClassListStr='';//document.all.ClassList.value;
	var NewsNumberStr='';
	if (document.all.NewsNumber.value=='') NewsNumberStr='10';
	else NewsNumberStr=document.all.NewsNumber.value;
	var RowNumberStr='';
	if (document.all.RowNumber.value=='') RowNumberStr='1';
	else RowNumberStr=document.all.RowNumber.value;
	var NaviPicStr=document.all.NaviPic.value;
	var BGPicStr=document.all.BGPic.value;
	var RowHeightStr='';
	if (document.all.RowHeight.value=='') RowHeightStr='20';
	else RowHeightStr=document.all.RowHeight.value;
	var CssFileStr=document.all.CssFile.value;
	var OpenModeStr=document.all.OpenMode.value;
	var DetachPageStr=document.all.DetachPage.value;
	var TitleNumberStr='';
	if (document.all.TitleNumber.value=='') TitleNumberStr='10';
	else TitleNumberStr=document.all.TitleNumber.value;
	var DateRuleStr=document.all.DateRule.value;
	var DateRightStr=document.all.DateRight.value;
	var DateCSSStyleStr=document.all.DateCSSStyle.value;
	var TxtNaviStr=document.all.TxtNavi.value;
	window.returnValue='{%=ClassNewsList("'+ClassListStr+'","'+NewsNumberStr+'","'+RowNumberStr+'","'+NaviPicStr+'","'+BGPicStr+'","'+RowHeightStr+'","'+CssFileStr+'","'+OpenModeStr+'","'+DetachPageStr+'","'+TitleNumberStr+'","'+DateRuleStr+'","'+DateRightStr+'","'+DateCSSStyleStr+'","'+TxtNaviStr+'")%}';
	window.close();
}
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
</script>