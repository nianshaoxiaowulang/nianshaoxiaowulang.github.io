<% Option Explicit %>
<!--#include file="Function.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P030300") then Call ReturnError1()
Dim TempClassListStr
	TempClassListStr = ClassList("ClassID")

%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��Ŀ��ҳ����</title>
</head>
<body topmargin="2" leftmargin="2" oncontextmenu="return false;">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="28" colspan="3" class="ButtonListLeft">
<div align="center"><strong>��Ŀ��ҳ���ɹ���</strong></div></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="8" cellpadding="0">
  <tr> 
    <td width="8%">&nbsp;</td>
    <td width="4%">&nbsp;</td>
    <td colspan="2">&nbsp;</td>
  </tr>
  <form action="RefreshClassSave.asp?Types=ClassAll" method="post" name="ReClassAllForm">
    <tr> 
      <td>&nbsp;</td>
      <td>ȫ��</td>
      <td colspan="2"><input name="imageField2" type="image" src="../../Images/Publish.gif" width="75" height="21" border="0"></td>
    </tr>
  </form>
  <form action="RefreshClassSave.asp?Types=ClassOne" method="post" name="ReClassOneForm">
    <tr> 
      <td rowspan="2">&nbsp;</td>
      <td height="53" valign="top">����</td>
      <td width="24%" rowspan="2" valign="top"> <p> 
          <select name="ClassID" size=13 multiple style="width:170">
            <% =TempClassListStr %>
          </select>
          <br>
          <input name="IssueSubClass" type="checkbox" id="IssueSubClass3" value="IssueSubClass">
          ��������Ŀ����������Ŀ <br>
          <input name="imageField22" type="image" src="../../Images/Publish.gif"  border="0">
      </td>
      <td width="64%" rowspan="2" valign="top"><p><font color=red>˵��:</font></p>
        <p><font color=red>1�������԰�סCTRL��SHIFTͬʱѡ������Ŀһ�𷢲�</font></p>
        <p><font color=red>2��Ҳ����ѡ��һ����Ŀ��Ȼ��ѡ���������Ŀ����</font></p>
        <p><font color=red>3����ѡ������Ŀ��ʱ�򣬰�������Ŀ����������</font></p>
        <p><font color=red>4�������Ҫ���ɵ���Ŀ�϶࣬������÷�������</font></p>
        <p><font color=red>5��ע�������ɹ����У������ֶ�ˢ�´�ҳ��</font></p></td>
    </tr>
    <tr> 
      <td valign="top">&nbsp;</td>
    </tr>
  </form>
  <tr> 
    <td>&nbsp;</td>
    <td colspan="3">&nbsp;</td>
  </tr>
</table>
</body>
</html>
