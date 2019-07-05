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
<title>选择专题标签属性</title>
<link rel="stylesheet" href="../../CSS/ModeWindow.css">
<script language="JavaScript" src="../SysJS/PublicJS.js"></script>
</head>
<body topmargin="0" leftmargin="0" scroll=no>
<div align="center"> 
  <table width="96%" border="0" cellpadding="0" cellspacing="0">
    <tr> 
      <td height="30" colspan="2">专题列表 
        <select  style="width:85%;" onChange="ChangeInsert(this);" name="SpecialList">
          <option selected value="">选择专题</option>
          <%
			Dim RsSpecialObj
			Set RsSpecialObj = Conn.Execute("Select EName,CName,SpecialID from FS_Special")
			do while Not RsSpecialObj.Eof
			%>
          <option value="<% = RsSpecialObj("EName") %>"> 
          <% = RsSpecialObj("CName") %>
          </option>
          <%
				RsSpecialObj.MoveNext
			loop
			Set RsSpecialObj = Nothing
			%>
        </select> <div align="left"></div></td>
    </tr>
    <tr> 
      <td width="50%" height="30">新闻数量 
        <input name="NewsNumber" type="text"  style="width:70%;" value="10"></td>
      <td height="30"><div align="left">标题字数 
          <input name="TitleNumber" type="text"  style="width:70%;" value="30">
        </div></td>
    </tr>
    <tr> 
      <td height="30" colspan="2">导航图片 
        <input type="text" readonly  style="width:63%;" id="NaviPic" name="NaviPic"> 
        <input name="Submitsdf" onClick="OpenWindowAndSetValue('SelectPic.asp?CurrPath=/<% = UpFiles %>&ShowVirtualPath=true',550,290,window,document.all.NaviPic);" type="button" id="Submitsdf" value="选择图片"> 
        <div align="left"></div></td>
    </tr>
    <tr> 
      <td height="30" colspan="2">分隔图片 
        <input type="text" readonly style="width:63%;" id="CompatPic" name="CompatPic"> 
        <input type="button" name="Submit3" onClick="OpenWindowAndSetValue('SelectPic.asp?CurrPath=/<% = UpFiles %>&ShowVirtualPath=true',550,290,window,document.all.CompatPic);" value="选择图片"> 
      </td>
    </tr>
    <tr> 
      <td height="30">文字导航 
        <input type="text" name="TxtNavi" style="width:70%;"></td>
      <td height="30">弹出窗口 
        <select  style="width:70%;" name="OpenType">
          <option value="0" selected>否</option>
          <option value="1">是</option>
        </select></td>
    </tr>
    <tr> 
      <td height="30">新闻行距 
        <input type="text"  style="width:70%;" value="20" name="RowHeight"></td>
      <td height="30"><div align="left">排列列数 
          <input type="text" onBlur="CheckNumber(this,'排列列数');"  style="width:70%;" value="1" name="RowNumber">
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
      <td height="30">标题样式 
        <input type="text" style="width:70%;" name="CSSStyle"> </td>
      <td height="30">日期样式 
        <input type="text" style="width:70%;" name="DateCSSStyle"></td>
    </tr>
    <tr> 
      <td height="30" colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td>&nbsp;</td>
            <td width="100"> <div align="center"> 
                <input name="SubmitBtn" disabled type="button" id="SubmitBtn" onClick="InsertScript();" value=" 确 定 ">
              </div></td>
            <td width="100"> <div align="center"> 
                <input type="button" onClick="window.close();" name="Submit2" value=" 取 消 ">
              </div></td>
            <td>&nbsp;</td>
          </tr>
        </table></td>
    </tr>
  </table>
</div>
</body>
</html>
<%
Set Conn = Nothing
%>
<script>
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
function ChangeInsert(Obj)
{
	var SpecialEName=Obj.value;
	if (SpecialEName!='')
	{
		document.all.SubmitBtn.disabled=false;
	}
	else
	{
		document.all.SubmitBtn.disabled=true;
	}
}
function InsertScript()
{
	var TempStr='';
	var NewsNumberStr='';
	var TitleNumberStr='';
	var CompatPicStr='';
	var NaviPicStr='';
	var DateRuleStr='';
	var DateRightStr='';
	var RowHeightStr='';
	var RowNumberStr='';
	var MoreLinkTypeStr=document.all.MoreLinkType.value;
	var MoreLinkContentStr=document.all.MoreLinkContent.value;
	var SpecialListObj=document.all.SpecialList(document.all.SpecialList.selectedIndex);
	
	if (document.all.NewsNumber.value=='') NewsNumberStr='10';
	else  NewsNumberStr=document.all.NewsNumber.value;
	
	if (document.all.TitleNumber.value=='') TitleNumberStr='10';
	else  TitleNumberStr=document.all.TitleNumber.value;
	
	if (document.all.CompatPic.value=='') CompatPicStr='';
	else  CompatPicStr=document.all.CompatPic.value;
	
	if (document.all.NaviPic.value=='') NaviPicStr='';
	else  NaviPicStr=document.all.NaviPic.value;
	
	if (document.all.DateRule.value=='') DateRuleStr='';
	else  DateRuleStr=document.all.DateRule.value;
	
	DateRightStr=document.all.DateRight.value;
	
	if (document.all.RowHeight.value=='') RowHeightStr='20';
	else  RowHeightStr=document.all.RowHeight.value;
	
	if (document.all.RowNumber.value=='') RowNumberStr='1';
	else  RowNumberStr=document.all.RowNumber.value;
	var CSSStyleStr=document.all.CSSStyle.value;
	var OpenTypeStr=document.all.OpenType.value;
	var DateCSSStyleStr=document.all.DateCSSStyle.value;
	var TxtNaviStr=document.all.TxtNavi.value;
	TempStr='{%=SpecialNewsList("'+SpecialListObj.value+'","'+NewsNumberStr+'","'+TitleNumberStr+'","'+CompatPicStr+'","'+NaviPicStr+'","'+DateRuleStr+'","'+DateRightStr+'","'+RowHeightStr+'","'+RowNumberStr+'","'+MoreLinkTypeStr+'","'+MoreLinkContentStr+'","'+CSSStyleStr+'","'+OpenTypeStr+'","'+DateCSSStyleStr+'","'+TxtNaviStr+'")%}';
	window.returnValue=TempStr;
	window.close();
}
</script>
