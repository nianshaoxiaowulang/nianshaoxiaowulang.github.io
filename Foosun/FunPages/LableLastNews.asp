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
<body topmargin="0" leftmargin="0" scroll=no>
<div align="center">
  <table width="96%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td height="30"> <div align="left">��Ŀ�б� 
          <select name="ClassList" id="ClassList" style="width:70%;">
            <option value="" selected>��Ŀѡ��</option>
            <% =TempClassListStr %>
          </select>
        </div></td>
	<td height="30"><div align="left">�������� 
        <select name="SoonClass" id="select" style="width:70%;">
          <option value="1" selected>��</option>
          <option value="0">��</option>
        </select>
     </div></td>
    </tr>
    <tr> 
      <td width="50%" height="30"> <div align="left">�������� 
          <input name="NewNumber" id="NewNumber" onBlur="CheckNumber(this,'��������');" type="text"  style="width:70%;" value="10">
        </div></td>
      <td>�������� 
        <input name="RowNumber" type="text" onBlur="CheckNumber(this,'��������');" id="RowNumber"  style="width:70%;" value="1"></td>
    </tr>
    <tr> 
      <td height="30"><div align="left">�������� 
          <input name="TitleNumber" id="TitleNumber" onBlur="CheckNumber(this,'��������');" type="text"  style="width:70%;" value="30">
        </div></td>
      <td>������ʽ
<input type="text"  style="width:70%;" name="CSSStyle"></td>
    </tr>
    <tr> 
      <td height="30" colspan="2"><div align="left">����ͼƬ 
          <input type="text" readonly  style="width:63%;" id="NaviPic" name="NaviPic">
          <input type="button" onClick="OpenWindowAndSetValue('SelectPic.asp?CurrPath=/<% = UpFiles %>&ShowVirtualPath=true',550,290,window,document.all.NaviPic);" name="Submit3" value="ѡ��ͼƬ">
        </div></td>
    </tr>
    <tr> 
      <td height="30" colspan="2"><div align="left">�ָ�ͼƬ 
          <input type="text" readonly  style="width:63%;" id="CompatPic2" name="CompatPic">
          <input type="button" onClick="OpenWindowAndSetValue('SelectPic.asp?CurrPath=/<% = UpFiles %>&ShowVirtualPath=true',550,290,window,document.all.CompatPic);" name="Submit4" value="ѡ��ͼƬ">
        </div></td>
    </tr>
    <tr> 
      <td height="30" colspan="2">���ֵ��� 
        <input type="text" name="TxtNavi" style="width:85%;"></td>
    </tr>
    <tr> 
      <td height="30"><div align="left">�������� 
          <select  style="width:70%;" name="OpenType">
            <option value="0" selected>��</option>
            <option value="1">��</option>
          </select>
        </div></td>
      <td>�����о� 
        <input name="RowHeight" type="text" style="width:70%;" id="RowHeight" value="20"></td>
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
	var NewNumberStr='';
	var TitleNumberStr='';
	var RowNumberStr='';
	if (document.all.NewNumber.value=='') NewNumberStr='10';
	else NewNumberStr=document.all.NewNumber.value;
	if (document.all.TitleNumber.value=='') TitleNumberStr='10';
	else TitleNumberStr=document.all.TitleNumber.value;
	if (document.all.RowNumber.value=='') RowNumberStr='10';
	else RowNumberStr=document.all.RowNumber.value;
	var NaviPicStr=document.all.NaviPic.value;
	var CompatPicStr=document.all.CompatPic.value;
	var OpenTypeStr=document.all.OpenType.value;
	var CSSStyleStr=document.all.CSSStyle.value;
	var RowHeightStr='';
	if (document.all.RowHeight.value=='') RowHeightStr='20';
	else RowHeightStr=document.all.RowHeight.value;
	var TxtNaviStr=document.all.TxtNavi.value;
	var ClassListStr=document.all.ClassList.value;
	var SoonClassStr=document.all.SoonClass.value;
	window.returnValue='{%=LastNews("'+ClassListStr+'","'+SoonClassStr+'","'+NewNumberStr+'","'+TitleNumberStr+'","'+RowNumberStr+'","'+NaviPicStr+'","'+CompatPicStr+'","'+OpenTypeStr+'","'+CSSStyleStr+'","'+RowHeightStr+'","'+TxtNaviStr+'")%}';
	window.close();
}
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
</script>