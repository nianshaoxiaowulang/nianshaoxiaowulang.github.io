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
<title>������������</title>
</head>
<link rel="stylesheet" href="../../CSS/ModeWindow.css">
<script language="JavaScript" src="../SysJS/PublicJS.js"></script>
<body topmargin="0" leftmargin="0" scroll=no scroll=no>
<div align="center">
  <table width="96%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td height="30"> <div align="left">��Ŀ�б� 
          <select name="ClassList" style="width:70%;">
            <option value="" selected>��Ŀѡ��</option>
            <% =TempClassListStr %>
          </select>
        </div></td>
    <td height="30"><div align="left">�������� 
        <select name="SoonClass" style="width:70%;">
          <option value="1" selected>��</option>
          <option value="0">��</option>
        </select>
      </div></td>
    </tr> 
	<tr> 
      <td width="50%" height="30"> <div align="left">ͼƬ��� 
          <input name="PicWidth" type="text" style="width:70%;" value="60">
        </div></td>
      <td height="30">ͼƬ�߶� 
        <input name="PicHeight" type="text" style="width:70%;" value="60"></td>
    </tr>
    <tr> 
      <td height="30" colspan="2"><div align="left">����ͼƬ 
          <input type="text" readonly  style="width:63%;" id="NaviPic" name="NaviPic">
          <input type="button" onClick="OpenWindowAndSetValue('SelectPic.asp?CurrPath=/<% = UpFiles %>&ShowVirtualPath=true',550,290,window,document.all.NaviPic);" name="Submit" value="ѡ��ͼƬ">
        </div></td>
    </tr>
    <tr> 
      <td height="30" colspan="2"><div align="left">�ָ�ͼƬ 
          <input type="text" readonly  style="width:63%;" id="BGPic" name="BGPic">
          <input type="button" onClick="OpenWindowAndSetValue('SelectPic.asp?CurrPath=/<% = UpFiles %>&ShowVirtualPath=true',550,290,window,document.all.BGPic);" name="Submit3" value="ѡ��ͼƬ">
        </div></td>
    </tr>
    <tr> 
      <td height="30" colspan="2">���ֵ��� 
        <input type="text" name="TxtNavi" style="width:85%;"></td>
    </tr>
    <tr> 
      <td height="30"><div align="left">�������� 
          <input name="TitleNumber" id="TitleNumber" onBlur="CheckNumber(this,'��������');" type="text"  style="width:70%;" value="30">
        </div></td>
      <td height="30">������ʽ
<input type="text"  style="width:70%;" name="CSSStyle"></td>
    </tr>
    <tr> 
      <td height="30"><div align="left">�������� 
          <input name="NewsNumber" id="NewsNumber" onBlur="CheckNumber(this,'��������');" type="text"  style="width:70%;" value="10">
        </div></td>
      <td height="30">�������� 
        <input name="RowNumber" type="text" onBlur="CheckNumber(this,'��������');" id="RowNumber"  style="width:70%;" value="1"></td>
    </tr>
    <tr> 
      <td height="30"><div align="left">�������� 
          <select  style="width:70%;" name="OpenType">
            <option value="0" selected>��</option>
            <option value="1">��</option>
          </select>
        </div></td>
      <td height="30"> �����о� 
        <input name="RowHeight" type="text" style="width:70%;" id="RowHeight" value="20"></td>
    </tr>
    <tr> 
      <td height="30"><div align="center"> 
          <input name="Submitafd" type="button" onClick="InsertScript();" id="Submitafd" value=" ȷ �� ">
        </div></td>
      <td height="30"><div align="center"> 
          <input type="button" onClick="window.close();" name="Submit2" value=" ȡ �� ">
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