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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>终极栏目下载标签属性</title>
</head>
<link rel="stylesheet" href="../../CSS/ModeWindow.css">
<script language="JavaScript" src="../SysJS/PublicJS.js"></script>
<body topmargin="0" leftmargin="0">
<div align="center">
  <table width="90%" border="0" cellspacing="0" cellpadding="0">
    <form action="" method="post" name="CNListForm">
      <tr> 
        <td height="30"><div align="left">排列列数 
            <input name="RowNumber" type="text" id="RowNumber" style="width:70%;" value="1">
          </div></td>
        <td height="30"><div align="left">标题字数 
            <input name="TitleNumber" type="text" id="TitleNumber" style="width:70%;" value="30">
          </div></td>
      </tr>
      <tr> 
        <td height="30" colspan="2"><div align="left">分隔图片 
            <input name="BGPic" type="text" readonly id="BGPic" style="width:63%;">
            <input type="button" name="sdafsdf" value="选择图片" onClick="OpenWindowAndSetValue('SelectPic.asp?CurrPath=/<% = UpFiles %>&ShowVirtualPath=true',550,290,window,document.CNListForm.BGPic);">
          </div></td>
      </tr>
      <tr> 
        <td height="30"><div align="left">下载行距 
            <input name="RowHeight" type="text" style="width:70%;" value="20">
          </div></td>
        <td height="30"><div align="left">标题样式
<input type="text" style="width:70%;" name="CssFile" id="CssFile">
          </div></td>
      </tr>
      <tr> 
        <td height="30"><div align="left">弹出窗口 
            <select name="OpenMode" id="OpenMode" style="width:70%">
              <option value="1">是</option>
              <option value="0" selected>否</option>
            </select>
          </div></td>
        <td height="30"><div align="left">是否分页 
            <select name="DetachPage" id="DetachPage" style="width:70%">
              <option value="1" selected>是</option>
              <option value="0">否</option>
            </select>
          </div></td>
      </tr>
      <tr> 
        <td height="30" colspan="2"><div align="left">列表样式 
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
            </select>
            <input name="Submitfasd" type="button" id="Submitfasd" onClick="BrowStyle();" value=" 查 看 ">
          </div></td>
      </tr>
      <tr> 
        <td width="42%" height="30">分页数量 
          <input name="NewsNumber" type="text" id="NewsNumber" style="width:70%;" value="10"> 
        </td>
        <td width="50%"><div align="left"></div></td>
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
	var NaviPicStr='';//document.all.NaviPic.value;
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
	var DownListStyleStr=document.all.DownListStyle.value;
	var TxtNaviStr='';//document.all.TxtNavi.value;
	window.returnValue='{%=DownLoadList("'+ClassListStr+'","'+NewsNumberStr+'","'+RowNumberStr+'","'+NaviPicStr+'","'+BGPicStr+'","'+RowHeightStr+'","'+CssFileStr+'","'+OpenModeStr+'","'+DetachPageStr+'","'+TitleNumberStr+'","'+DownListStyleStr+'","'+TxtNaviStr+'")%}';
	window.close();
}
function BrowStyle()
{
	if (document.CNListForm.DownListStyle.value!='') OpenWindow('Templet_DownStyleBrowFrame.asp?FileName=Templet_DownStyleBrow.asp&ID='+document.all.DownListStyle.value,360,190,window);
}
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
</script>