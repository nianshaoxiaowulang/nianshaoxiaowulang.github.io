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
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��Ŀ�����б�����</title>
</head>
<link rel="stylesheet" href="../../CSS/ModeWindow.css">
<script language="JavaScript" src="../SysJS/PublicJS.js"></script>
<body topmargin="0" leftmargin="0">
<div align="center">
  <table width="90%" border="0" cellspacing="0" cellpadding="0">
    <form action="" method="post" name="CNListForm">
      <tr> 
        <td width="42%" height="30">��ҳ���� 
          <input name="NewsNumber" type="text" id="NewsNumber2" style="width:70%;" value="10"> 
        </td>
        <td width="50%"><div align="left">�����о� 
            <input name="RowHeight" type="text" style="width:70%;" value="20">
          </div></td>
      </tr>
      <tr> 
        <td height="30"><div align="left">�������� 
            <input name="RowNumber" type="text" id="RowNumber" style="width:70%;" value="1">
          </div></td>
        <td height="30"><div align="left">�������� 
            <input name="TitleNumber" type="text" id="TitleNumber" style="width:70%;" value="40">
          </div></td>
      </tr>
      <tr> 
        <td height="30" colspan="2"><div align="left">����ͼƬ
            <input name="NaviPic" readonly type="text" id="NaviPic" style="width:63%;">
            <input type="button" name="Subfdgf" value="ѡ��ͼƬ" onClick="OpenWindowAndSetValue('SelectPic.asp?CurrPath=/<% = UpFiles %>&ShowVirtualPath=true',550,290,window,document.CNListForm.NaviPic);">
          </div></td>
      </tr>
      <tr> 
        <td height="30" colspan="2"><div align="left">�ָ�ͼƬ 
            <input name="BGPic" readonly type="text" id="BGPic" style="width:63%;">
            <input type="button" name="sdafsdf" value="ѡ��ͼƬ" onClick="OpenWindowAndSetValue('SelectPic.asp?CurrPath=/<% = UpFiles %>&ShowVirtualPath=true',550,290,window,document.CNListForm.BGPic);">
          </div></td>
      </tr>
      <tr> 
        <td height="30" colspan="2">���ֵ��� 
          <input type="text" name="TxtNavi" style="width:85%;"></td>
      </tr>
      <tr> 
        <td height="30"><div align="left">������ʽ 
            <input type="text" style="width:70%;" name="CssFile" id="CssFile">
          </div></td>
        <td height="30"><div align="left">������ʽ 
            <input type="text" style="width:70%;" name="DateCSSStyle">
          </div></td>
      </tr>
      <tr> 
        <td height="30"><div align="left">���ڸ�ʽ 
            <select  style="width:70%;" name="DateRule" id="DateRule">
              <option selected>ѡ�����ڸ�ʽ</option>
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
              <option value="11">9��1</option>
              <option value="12">1��11ʱ</option>
              <option value="13">1��11��</option>
              <option value="14">11ʱ11��</option>
              <option value="15">11:11</option>
              <option value="16">2004��9��1��</option>
            </select>
          </div></td>
        <td height="30"><div align="left">���ڶ��� 
            <select  style="width:70%;" name="DateRight">
              <option value="Right">�Ҷ���</option>
              <option value="Left" selected>�����</option>
              <option value="Center">����</option>
            </select>
          </div></td>
      </tr>
      <tr> 
        <td height="30"><div align="left">�������� 
            <select name="OpenMode" id="OpenMode" style="width:70%">
              <option value="1">��</option>
              <option value="0" selected>��</option>
            </select>
          </div></td>
        <td height="30"><div align="left">���ŷ�ҳ 
            <select name="DetachPage" id="DetachPage" style="width:70%">
              <option value="1" selected>��</option>
              <option value="0">��</option>
            </select>
          </div></td>
      </tr>
      <tr> 
        <td height="30" colspan="2"><div align="center"> 
            <input type="button" onClick="InsertScript();" name="Submit" value=" ȷ �� ">
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
            <input type="button" onClick="window.close();" name="Submit2" value=" ȡ �� ">
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
	var NaviPicStr=document.all.NaviPic.value;
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
	var DateRuleStr=document.all.DateRule.value;
	var DateRightStr=document.all.DateRight.value;
	var DateCSSStyleStr=document.all.DateCSSStyle.value;
	var TxtNaviStr=document.all.TxtNavi.value;
	window.returnValue='{%=ClassNewsList("'+ClassListStr+'","'+NewsNumberStr+'","'+RowNumberStr+'","'+NaviPicStr+'","'+BGPicStr+'","'+RowHeightStr+'","'+CssFileStr+'","'+OpenModeStr+'","'+DetachPageStr+'","'+TitleNumberStr+'","'+DateRuleStr+'","'+DateRightStr+'","'+DateCSSStyleStr+'","'+TxtNaviStr+'")%}';
	window.close();
}
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
</script>