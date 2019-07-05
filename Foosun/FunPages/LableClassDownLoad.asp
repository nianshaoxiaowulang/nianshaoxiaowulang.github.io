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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>栏目下载标签属性</title>
</head>
<link rel="stylesheet" href="../../CSS/ModeWindow.css">
<script language="JavaScript" src="../SysJS/PublicJS.js"></script>
<body topmargin="0" leftmargin="0">
<div align="center">
  <table width="96%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="50%" height="30">栏目列表 
        <select onChange="ChangeInsertState(this);" name="ClassList" id="ClassList" style="width:70%;">
          <option value="" selected>栏目选择</option>
          <% =TempClassListStr %>
        </select> </td>
      <td>下载行距 
        <input type="text" style="width:70%;" value="20" name="RowHeight"> </td>
    </tr>
    <tr> 
      <td height="30"> 调用数量 
        <input name="NewsListNumber" onBlur="CheckNumber(this,'新闻数量');" type="text"    style="width:70%;" value="10"> 
      </td>
      <td>标题字数 
        <input name="TitleNumber" onBlur="CheckNumber(this,'标题字数');" type="text"    style="width:70%;" value="30"> 
      </td>
    </tr>
    <tr> 
      <td height="30" colspan="2"><div align="left">分隔图片 
          <input type="text" readonly style="width:63%;" id="CompatPic" name="CompatPic">
          <input type="button" onClick="OpenWindowAndSetValue('SelectPic.asp?CurrPath=/<% = UpFiles %>&ShowVirtualPath=true',550,290,window,document.all.CompatPic);" name="Submit" value="选择图片">
        </div></td>
    </tr>
    <tr> 
      <td height="30">排列列数 
        <input type="text" onBlur="CheckNumber(this,'排列列数');"  style="width:70%;" value="1" name="RowNumber"> 
      </td>
      <td height="30">日期格式 
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
        </select></td>
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
      <td height="30">弹出窗口 
        <select  style="width:70%;" name="OpenType">
          <option value="0" selected>否</option>
          <option value="1">是</option>
        </select> </td>
      <td height="30" nowrap>标题样式
<input type="text" style="width:70%;" name="CSSStyle"></td>
    </tr>
    <tr> 
      <td height="30" colspan="2">列表样式 
        <select name="DownListStyle" style="width:65%;">
          <%
		Dim StyleSql,RsStyleObj
		StyleSql = "Select * from FS_DownListStyle"
		Set RsStyleObj = Conn.Execute(StyleSql)
		do while Not RsStyleObj.Eof
		%>
          <option value="<% = RsStyleObj("ID") %>"> 
          <% = RsStyleObj("Name") %>
          </option>
          <%
			RsStyleObj.MoveNext
		loop
		Set RsStyleObj = Nothing
		%>
        </select> <input name="Submitfasd" type="button" id="Submitfasd" onClick="BrowStyle();" value=" 查 看 "></td>
    </tr>
    <tr> 
      <td height="30" colspan="2"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td>&nbsp;</td>
            <td width="100"> <div align="center"> 
                <input name="SubmitBtn" type="button" disabled id="Submitsss4" onClick="InsertScriptFun();" value=" 确 定 ">
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
function ChangeInsertState(Obj)
{
	var ClassEName=Obj.options(Obj.selectedIndex).value;
	if (ClassEName!='')
	{
		document.all.SubmitBtn.disabled=false;
	}
	else
	{
		document.all.SubmitBtn.disabled=true;
	}
}
function InsertScriptFun(Obj)
{
	var TempStr='';
	var NewsListNumberStr='';
	var TitleNumberStr='';
	var CompatPicStr='';
	var NaviPicStr='';
	var DateRuleStr='';
	var DateRightStr='';
	var RowHeightStr='';
	var RowNumberStr='';
	var MoreLinkTypeStr=document.all.MoreLinkType.value;
	var MoreLinkContentStr=document.all.MoreLinkContent.value;
	var ClassListObj=document.all.ClassList.options(document.all.ClassList.selectedIndex);
	if (document.all.NewsListNumber.value=='') NewsListNumberStr='10';
	else  NewsListNumberStr=document.all.NewsListNumber.value;
	if (document.all.TitleNumber.value=='') TitleNumberStr='10';
	else  TitleNumberStr=document.all.TitleNumber.value;
	CompatPicStr=document.all.CompatPic.value;
	//if (document.all.NaviPic.value=='') NaviPicStr='';
	//else  NaviPicStr=document.all.NaviPic.value;
	DateRuleStr=document.all.DateRule.value;
	//DateRightStr=document.all.DateRight.value;
	
	if (document.all.RowHeight.value=='') RowHeightStr='20';
	else  RowHeightStr=document.all.RowHeight.value;
	
	if (document.all.RowNumber.value=='') RowNumberStr='1';
	else  RowNumberStr=document.all.RowNumber.value;
	var OpenTypeStr=document.all.OpenType.value;
	var CSSStyleStr=document.all.CSSStyle.value;
	var ShowClassCNNameStr='';
	//ShowClassCNNameStr=document.all.ShowClassCNName.value;
	var DownListStyleStr=document.all.DownListStyle.value;
	var TxtNaviStr='';
	TempStr='{%=ClassDownLoad("'+ClassListObj.value+'","'+NewsListNumberStr+'","'+TitleNumberStr+'","'+CompatPicStr+'","'+NaviPicStr+'","'+DateRuleStr+'","'+DateRightStr+'","'+RowHeightStr+'","'+RowNumberStr+'","'+ShowClassCNNameStr+'","'+MoreLinkTypeStr+'","'+MoreLinkContentStr+'","'+CSSStyleStr+'","'+OpenTypeStr+'","'+DownListStyleStr+'","'+TxtNaviStr+'")%}';
	window.returnValue=TempStr;
	window.close();
}
function BrowStyle()
{
	if (document.all.DownListStyle.value!='') OpenWindow('Templet_DownStyleBrowFrame.asp?FileName=Templet_DownStyleBrow.asp&ID='+document.all.DownListStyle.value,360,190,window);
}
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
</script>
