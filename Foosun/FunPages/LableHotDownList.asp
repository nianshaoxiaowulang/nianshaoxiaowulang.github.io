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
<title>热点下载标签属性</title>
</head>
<link rel="stylesheet" href="../../CSS/ModeWindow.css">
<script language="JavaScript" src="../SysJS/PublicJS.js"></script>
<body topmargin="0" leftmargin="0">
<div align="center">
  <table width="96%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td height="30" colspan="2"><div align="left">栏目列表 
          <select name="ClassList" id="ClassList" style="width:85%;">
            <option value="" selected>栏目选择</option>
            <% =TempClassListStr %>
          </select>
        </div></td>
    </tr>
    <tr> 
      <td width="50%" height="30"> <div align="left">新闻数量 
          <input name="NewNumber" id="NewNumber" onBlur="CheckNumber(this,'新闻数量');" type="text"  style="width:70%;" value="10">
        </div></td>
      <td>新闻列数 
        <input name="RowNumber" type="text" onBlur="CheckNumber(this,'新闻列数');" id="RowNumber"  style="width:70%;" value="1"> 
      </td>
    </tr>
    <tr> 
      <td height="30"><div align="left">标题字数 
          <input name="TitleNumber" id="TitleNumber" onBlur="CheckNumber(this,'标题字数');" type="text"  style="width:70%;" value="30">
        </div></td>
      <td>标题样式
<input type="text"  style="width:70%;" name="CSSStyle"></td>
    </tr>
    <tr> 
      <td height="30" colspan="2"><div align="left">分隔图片 
          <input type="text" readonly  style="width:63%;" id="CompatPic2" name="CompatPic">
          <input type="button" onClick="OpenWindowAndSetValue('SelectPic.asp?CurrPath=/<% = UpFiles %>&ShowVirtualPath=true',550,290,window,document.all.CompatPic);" name="Submit3" value="选择图片">
        </div></td>
    </tr>
    <tr> 
      <td height="30"><div align="left">弹出窗口 
          <select  style="width:70%;" name="OpenType">
            <option value="0" selected>否</option>
            <option value="1">是</option>
          </select>
        </div></td>
      <td>新闻行距 
        <input name="RowHeight" type="text" style="width:70%;" id="RowHeight" value="20"></td>
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
      <td height="30" colspan="2"> <div align="right"> 
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td>&nbsp;</td>
              <td width="100"> <div align="center"> 
                  <input name="SubmitBtn" type="button" id="Submitsss4" onClick="InsertScript();" value=" 确 定 ">
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
	var NewNumberStr='';
	var TitleNumberStr='';
	var RowNumberStr='';
	if (document.all.NewNumber.value=='') NewNumberStr='10';
	else NewNumberStr=document.all.NewNumber.value;
	if (document.all.TitleNumber.value=='') TitleNumberStr='10';
	else TitleNumberStr=document.all.TitleNumber.value;
	if (document.all.RowNumber.value=='') RowNumberStr='10';
	else RowNumberStr=document.all.RowNumber.value;
	var NaviPicStr='';//document.all.NaviPic.value;
	var CompatPicStr=document.all.CompatPic.value;
	var OpenTypeStr=document.all.OpenType.value;
	var CSSStyleStr=document.all.CSSStyle.value;
	var RowHeightStr='';
	if (document.all.RowHeight.value=='') RowHeightStr='20';
	else RowHeightStr=document.all.RowHeight.value;
	var ClassListObj=document.all.ClassList.options(document.all.ClassList.selectedIndex);
	var DownListStyleStr=document.all.DownListStyle.value;
	var TxtNaviStr='';//document.all.TxtNavi.value;
	window.returnValue='{%=HotDownList("'+ClassListObj.value+'","'+NewNumberStr+'","'+TitleNumberStr+'","'+RowNumberStr+'","'+NaviPicStr+'","'+CompatPicStr+'","'+OpenTypeStr+'","'+CSSStyleStr+'","'+RowHeightStr+'","'+DownListStyleStr+'","'+TxtNaviStr+'")%}';
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
