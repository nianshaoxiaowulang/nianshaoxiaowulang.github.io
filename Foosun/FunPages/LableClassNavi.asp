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
<title>��Ŀ��������ѡ��</title>
</head>
<link rel="stylesheet" href="../../CSS/ModeWindow.css">
<script language="JavaScript" src="../SysJS/PublicJS.js"></script>
<body topmargin="0" leftmargin="0" scroll="no">
<div align="center">
  <table width="96%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="50%" height="30">�������� 
        <input name="RowNumber" type="text" style="width:70%;" value="10"></td>
      <td height="30">������ʽ 
        <input type="text" style="width:70%;" name="CSSStyle"> 
      </td>
    </tr>
    <tr> 
      <td height="30" colspan="2">����ͼƬ 
        <input type="text" readonly  style="width:63%;" name="NaviPic">
        <input type="button" onClick="OpenWindowAndSetValue('SelectPic.asp?CurrPath=/<% = UpFiles %>&ShowVirtualPath=true',550,290,window,document.all.NaviPic);" name="Submit" value="ѡ��ͼƬ"></td>
    </tr>
    <tr> 
      <td height="30" colspan="2">�ָ�ͼƬ 
        <input type="text" readonly style="width:63%;" name="CompatPic">
        <input type="button" name="Submit3" onClick="OpenWindowAndSetValue('SelectPic.asp?CurrPath=/<% = UpFiles %>&ShowVirtualPath=true',550,290,window,document.all.CompatPic);" value="ѡ��ͼƬ"></td>
    </tr>
    <tr> 
      <td height="30">���ֵ��� 
        <input type="text" name="TxtNavi" style="width:70%;"></td>
      <td height="30">�������� 
        <select style="width:70%;" name="OpenType">
          <option value="0" selected>��</option>
          <option value="1">��</option>
        </select></td>
    </tr>
    <tr> 
      <td height="30"><div align="center"> 
          <input name="Submitsdfds" onClick="InsertScript();" type="button" id="Submitsdfds" value=" ȷ �� ">
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
	var NaviPicStr='';
	var CompatPicStr='';
	var RowNumberStr='';
	var OpenTypeStr='';
	var CSSStyleStr='';
	NaviPicStr=document.all.NaviPic.value;
	CompatPicStr = document.all.CompatPic.value;
	if (document.all.RowNumber.value=='') RowNumberStr='10';
	else RowNumberStr=document.all.RowNumber.value;
	OpenTypeStr=document.all.OpenType.value;
	CSSStyleStr=document.all.CSSStyle.value;
	var TxtNaviStr=document.all.TxtNavi.value;
	window.returnValue='{%=ClassNavi("'+NaviPicStr+'","'+CompatPicStr+'","'+RowNumberStr+'","'+OpenTypeStr+'","'+CSSStyleStr+'","'+TxtNaviStr+'")%}';
	window.close();
}
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
</script>