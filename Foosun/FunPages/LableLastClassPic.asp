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
<link rel="stylesheet" href="../../CSS/ModeWindow.css">
<script language="JavaScript" src="../SysJS/PublicJS.js"></script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ռ�ͼƬ</title>
</head>
<body>
<table width="96%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="42%" height="30"> <div align="left">��ҳ���� 
        <input name="NewsNumber" type="text" id="NewsNumber" style="width:70%;" value="10">
      </div></td>
    <td width="50%"><div align="left">�Ƿ��ҳ 
        <select name="CutPage" id="CutPage" style="width:70%;">
          <option value="1" selected>��</option>
          <option value="0">��</option>
        </select>
      </div></td>
  </tr>
  <tr> 
    <td height="30"><div align="left">��ʾ���� 
        <select name="ShowTitle" id="ShowTitle" style="width:70%;">
          <option value="1" selected>��</option>
          <option value="0">��</option>
        </select>
      </div></td>
    <td height="30"><div align="left">�������� 
        <select name="OpenMode" id="OpenMode" style="width:70%;">
          <option value="1">��</option>
          <option value="0" selected>��</option>
        </select>
      </div></td>
  </tr>
  <tr> 
    <td height="30"><div align="left">�������� 
        <input name="TitleNumber" type="text" id="TitleNumber" style="width:70%;" value="30">
      </div></td>
    <td height="30"><div align="left">�������� 
        <input name="RowNum" type="text" id="RowNum" style="width:70%;" value="1">
      </div></td>
  </tr>
  <tr> 
    <td height="30"><div align="left">ͼƬ��� 
        <input name="PicWidth" type="text" id="PicWidth" style="width:70%;" value="60">
      </div></td>
    <td height="30"><div align="left">ͼƬ�߶� 
        <input name="PicHeight" type="text" id="PicHeight2" style="width:70%;" value="60">
      </div></td>
  </tr>
  <tr> 
    <td height="30"><div align="left">�䡡���� 
        <input name="RowSpace" type="text" id="RowSpace" style="width:70%;" value="6">
      </div></td>
    <td height="30"><div align="left">������ʽ
<input type="text" style="width:70%;" name="CssFile" id="CssFile">
      </div></td>
  </tr>
  <tr> 
    <td height="30" colspan="2"><div align="center"> 
        <input type="button" onClick="InsertScript();" name="Submit" value=" ȷ �� ">
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="button" onClick="window.close();" name="Submit2" value=" ȡ �� ">
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
	var CutPageStr=document.all.CutPage.value;
	var ShowTitleStr=document.all.ShowTitle.value;
	var OpenModeStr=document.all.OpenMode.value;
	var TitleNumberStr='';
	if (document.all.TitleNumber.value=='') TitleNumberStr='20';
	else TitleNumberStr=document.all.TitleNumber.value;
	var RowNumStr='';
	if (document.all.RowNum.value=='') RowNumStr='1';
	else RowNumStr=document.all.RowNum.value;
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
	window.returnValue='{%=LastClassPic("'+CutPageStr+'","'+NewsNumberStr+'","'+ShowTitleStr+'","'+TitleNumberStr+'","'+OpenModeStr+'","'+RowNumStr+'","'+PicWidthStr+'","'+PicHeightStr+'","'+CssFileStr+'","'+RowSpaceStr+'")%}';
	window.close();
}
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
</script>