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
<title>��������ѡ��</title>
</head>
<link rel="stylesheet" href="../../CSS/ModeWindow.css">
<script language="JavaScript" src="../SysJS/PublicJS.js"></script>
<body topmargin="0" leftmargin="0">
<div align="center">
  <table width="96%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="50%" height="30" nowrap> <div align="left">�������� 
          <select name="NaviType" id="NaviType"  style="width:70%;">
            <option value="1" selected>��Ŀ����</option>
            <option value="2">ר�⵼��</option>
            <option value="3">�������</option>
            <option value="4">��Ŀ+ר�⵼��</option>
            <option value="5">��Ŀ+�������</option>
            <option value="6">ר��+�������</option>
            <option value="7">��Ŀ+ר��+�������</option>
          </select>
        </div></td>
      <td width="50%" nowrap>�������� 
        <input name="RowNumber" onBlur="CheckNumber(this,'��ʾ����');" type="text"  style="width:70%;" value="10" size="6"> 
      </td>
    </tr>
    <tr> 
      <td height="30" colspan="2"><div align="left">����ͼƬ 
          <input type="text" readonly  style="width:63%;" id="NaviPic" name="NaviPic">
          <input type="button" name="Submit3" onClick="OpenWindowAndSetValue('SelectPic.asp?CurrPath=/<% = UpFiles %>&ShowVirtualPath=true',550,290,window,document.all.NaviPic);" value="ѡ��ͼƬ">
        </div></td>
    </tr>
    <tr> 
      <td height="30" colspan="2">�ָ�ͼƬ 
        <input type="text"  style="width:63%;" readonly name="BGPic"> <input type="button" name="Submit4" onClick="OpenWindowAndSetValue('SelectPic.asp?CurrPath=/<% = UpFiles %>&ShowVirtualPath=true',550,290,window,document.all.BGPic);" value="ѡ��ͼƬ"></td>
    </tr>
    <tr> 
      <td height="30" colspan="2">���ֵ��� 
        <input type="text" name="TxtNavi" style="width:85%;"></td>
    </tr>
    <tr> 
      <td height="30"><div align="left">�������� 
          <select style="width:70%;" name="OpenType">
            <option value="0" selected>��</option>
            <option value="1">��</option>
          </select>
        </div></td>
      <td>������ʽ
<input type="text" style="width:70%;" name="CSSStyle"></td>
    </tr>
    <tr> 
      <td height="30"><div align="center"> 
          <input type="button" name="Submit" onClick="InsertScript();" value=" ȷ �� ">
        </div></td>
      <td><div align="center"> 
          <input type="button" name="Submit2" onClick="window.close();" value=" ȡ �� ">
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