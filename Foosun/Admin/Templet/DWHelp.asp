<% Option Explicit %>
<!--#include file="../../../Inc/NoSqlhack.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P080800") then Call ReturnError1()
dim MallConfig
Set MallConfig=conn.execute("Select IsShop from FS_Config")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>DreamWeaver�����������</title>
</head>

<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<body topmargin="2" leftmargin="2">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle"> <div align="center"><strong>DreamWeaver�������</strong></div></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><font color="#0000FF"> һ���˹��ܸ����û���Dreamweaver��ʹ�÷�Ѷ�ṩ����չ�����дӢ�����ƣ�������ʽ����Ʒ�б���ʽ��<br>
      ��������ʹ���ж���ʹ�û�����㣬����ƶ�����Ӧ���ı������棬����ճ����DreamWeaver�������Ӧ���С� </font> </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="3">
  <tr> 
    <td height="30" colspan="4"> <div align="left"><strong><font color="#FF0000"> 
        ��ĿӢ�����ƶ��ձ�</font></strong></div>
      </td>
  </tr>
  <tr> 
    <td width="25%" height="30" class="ButtonListLeft"> <div align="center">��������</div></td>
    <td width="25%" class="ButtonList"> <div align="center">Ӣ������</div></td>
    <td width="25%" class="ButtonList"> <div align="center">��������</div></td>
    <td class="ButtonList"> <div align="center">Ӣ������</div></td>
  </tr>
<%
Dim RsClassObj,ClassSql,i
ClassSql = "Select * from FS_NewsClass"
Set RsClassObj = Conn.Execute(ClassSql)
i = 1
do while Not RsClassObj.Eof
%>
  <tr> 
    <td height="20">��<% = i & "��" %><% = RsClassObj("ClassCName") %></td>
    <td><input name="textfield" type="text" style="width:100%;" onClick="this.focus();this.select();" onMouseOver="this.focus();this.select();" onMouseOut="" value="<% = RsClassObj("ClassEName") %>" readonly></td>
<%
	RsClassObj.MoveNext
	i = i + 1
	if 	RsClassObj.Eof then Exit Do
%>
    <td>��<% = i & "��" %><% = RsClassObj("ClassCName") %></td>
    <td><input name="textfield" type="text" style="width:100%;" onClick="this.focus();this.select();" onMouseOver="this.focus();this.select();" onMouseOut="" value="<% = RsClassObj("ClassEName") %>" readonly></td>
  </tr>
<%
	RsClassObj.MoveNext
	i = i + 1
Loop
Set RsClassObj = Nothing
%>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="3">
  <tr> 
    <td height="30" colspan="4">
<div align="center"> 
        <div align="left"><font color="#FF0000"><strong>�����б���ʽID��Ӧ��</strong></font></div>
      </div></td>
  </tr>
  <tr> 
    <td width="50%" height="30" class="ButtonListLeft"> <div align="center">��ʽ����</div></td>
    <td width="25%" class="ButtonList">
<div align="center">ID</div></td>
    <td class="ButtonList">
<div align="center">��ʽ����</div></td>
  </tr>
<%
Dim RsDownStyleObj,DownStyleSql
DownStyleSql = "Select * from FS_DownListStyle"
Set RsDownStyleObj = Conn.Execute(DownStyleSql)
i = 1
do while Not RsDownStyleObj.Eof
%>
  <tr> 
    <td height="20">��<% = i & "��" %><% = RsDownStyleObj("Name") %></td>
    <td><input name="textfield" type="text" style="width:100%;" onClick="this.focus();this.select();" onMouseOver="this.focus();this.select();" onMouseOut="" value="<% = RsDownStyleObj("ID") %>" readonly></td>
    <td><div align="center"><span style="cursor:hand;" onClick="BrowStyle('Frame.asp?FileName=Templet_DownStyleBrow.asp&PageTitle=�鿴�����б���ʽ&ID=<% = RsDownStyleObj("ID") %>');">�鿴</span></div></td>
  </tr>
<%
	RsDownStyleObj.MoveNext
	i = i + 1
Loop
Set RsDownStyleObj = Nothing
%>
</table>
<%
If Cint(MallConfig(0))=1 then 
%>
<table width="100%" border="0" cellpadding="3" cellspacing="0" bgcolor="#CCCCCC">
  <tr bgcolor="#FFFFFF"> 
    <td height="30" colspan="4"> <div align="left"><strong><font color="#FF0000">��Ʒ�б���ʽID��Ӧ��</font></strong></div></td>
  </tr>
  <tr> 
    <td width="51%" height="30" class="ButtonListleft"> <div align="center">��ʽ����</div></td>
    <td width="25%" class="ButtonList">
<div align="center">ID</div></td>
    <td class="ButtonList">
<div align="center">��ʽ����</div></td>
  </tr>
<%
Dim RsMallObj,MallSql
MallSql = "Select * from FS_MallListStyle"
Set RsMallObj = Conn.Execute(MallSql)
i = 1
do while Not RsMallObj.Eof
%>
  <tr> 
    <td height="20" bgcolor="#FFFFFF"> 
      ��<% = i & "��" %><% = RsMallObj("Name") %></td>
    <td bgcolor="#FFFFFF"> 
      <input name="textfield" type="text" style="width:100%;" onClick="this.focus();this.select();" onMouseOver="this.focus();this.select();" onMouseOut="" value="<% = RsMallObj("ID") %>" readonly></td>
    <td bgcolor="#FFFFFF"> 
      <div align="center"><span style="cursor:hand;" onClick="BrowStyle('Frame.asp?FileName=Templet_DownStyleBrow.asp&PageTitle=�鿴��Ʒ�б���ʽ&ID=<% = RsMallObj("ID") %>');">�鿴</span></div></td>
  </tr>
<%
	RsMallObj.MoveNext
	i = i + 1
Loop
Set RsMallObj = Nothing
%>
</table>
<%
End IF%>
</body>
</html>
<%
Set Conn = Nothing
%>
<script language="JavaScript">
function BrowStyle(URL)
{
	OpenWindow(URL,360,190,window);
}
</script>