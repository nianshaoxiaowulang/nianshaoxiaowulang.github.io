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
<html>
<head>
<link rel="stylesheet" href="../Inc/ModeWindow.css">
<link rel="stylesheet" href="../../CSS/ModeWindow.css">
<script language="JavaScript" src="../SysJS/PublicJS.js"></script>
<title>���ʻع�</title>
</head>
<body>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="42%" height="30"> <div align="left">�������� 
        <input name="NewsNumber" type="text" id="NewsNumber" style="width:70%;" value="10">
      </div></td>
    <td width="50%"><div align="left">��Ŀ�б� 
        <select name="ClassList" id="ClassList" style="width:70%;">
          <option value="" selected>��Ŀѡ��</option>
          <% =TempClassListStr %>
        </select>
      </div></td>
  </tr>
  <tr> 
    <td height="30"><div align="left">�������� 
        <input name="TitleNumber" type="text" id="TitleNumber" style="width:70%;" value="30">
      </div></td>
    <td height="30"><div align="left">�������� 
        <select name="OpenMode" id="OpenMode" style="width:70%;">
          <option value="1" selected>��</option>
          <option value="0">��</option>
        </select>
      </div></td>
  </tr>
  <tr> 
    <td height="30"><div align="left">�����о� 
        <input name="RowSpace" type="text" id="RowSpace2" style="width:70%;" value="20">
      </div></td>
    <td height="30"><div align="left">�������� 
        <input name="RowNum" type="text" id="RowNum" style="width:70%;" value="1">
      </div></td>
  </tr>
  <tr> 
    <td height="30" colspan="2"><div align="left">����ͼƬ 
        <input type="text" readonly  style="width:63%;" id="NaviPic" name="NaviPic">
        <input type="button" name="Submit3" onClick="OpenWindowAndSetValue('SelectPic.asp?CurrPath=/<% = UpFiles %>&ShowVirtualPath=true',550,290,window,document.all.NaviPic);" value="ѡ��ͼƬ">
      </div></td>
  </tr>
  <tr> 
    <td height="30" colspan="2"><div align="left">�ָ�ͼƬ 
        <input type="text" readonly  style="width:63%;" id="CompatPic" name="CompatPic">
        <input type="button" onClick="OpenWindowAndSetValue('SelectPic.asp?CurrPath=/<% = UpFiles %>&ShowVirtualPath=true',550,290,window,document.all.CompatPic);" name="Submit4" value="ѡ��ͼƬ">
      </div></td>
  </tr>
  <tr> 
    <td height="30" colspan="2">���ֵ��� 
      <input type="text" name="TxtNavi" style="width:85%;"></td>
  </tr>
  <tr> 
    <td height="30"><div align="left">�������� 
        <select name="SoonClass" id="select" style="width:70%;">
          <option value="1" selected>��</option>
          <option value="0">��</option>
        </select>
      </div></td>
    <td height="30"><div align="left">������ʽ 
        <input type="text" style="width:70%;" name="CssFile" id="CssFile2">
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
	var ClassListStr=document.all.ClassList.value;
	var OpenModeStr=document.all.OpenMode.value;
	var TitleNumberStr='';
	if (document.all.TitleNumber.value=='') TitleNumberStr='20';
	else TitleNumberStr=document.all.TitleNumber.value;
	var RowNumStr='';
	if (document.all.RowNum.value=='') RowNumStr='1';
	else RowNumStr=document.all.RowNum.value;
	var SoonClassStr=document.all.SoonClass.value;
	var NaviPicStr=document.all.NaviPic.value;
	var CompatPicStr=document.all.CompatPic.value;
	var CssFileStr=document.all.CssFile.value;
	var RowSpaceStr='';
	if (document.all.RowSpace.value=='') RowSpaceStr='20';
	else RowSpaceStr=document.all.RowSpace.value;
	var TxtNaviStr=document.all.TxtNavi.value;
	window.returnValue='{%=TodayNews("'+ClassListStr+'","'+NewsNumberStr+'","'+SoonClassStr+'","'+TitleNumberStr+'","'+RowNumStr+'","'+NaviPicStr+'","'+CompatPicStr+'","'+OpenModeStr+'","'+CssFileStr+'","'+RowSpaceStr+'","'+TxtNaviStr+'")%}';
	window.close();
}
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
</script>