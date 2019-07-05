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
<title>DreamWeaver插件辅助工具</title>
</head>

<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<body topmargin="2" leftmargin="2">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle"> <div align="center"><strong>DreamWeaver插件帮助</strong></div></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><font color="#0000FF"> 一、此功能辅助用户在Dreamweaver中使用风讯提供的扩展插件书写英文名称，下载样式，商品列表样式。<br>
      二、请在使用中对照使用会更方便，鼠标移动到相应的文本框上面，复制粘贴到DreamWeaver插件的相应表单中。 </font> </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="3">
  <tr> 
    <td height="30" colspan="4"> <div align="left"><strong><font color="#FF0000"> 
        栏目英文名称对照表</font></strong></div>
      </td>
  </tr>
  <tr> 
    <td width="25%" height="30" class="ButtonListLeft"> <div align="center">中文名称</div></td>
    <td width="25%" class="ButtonList"> <div align="center">英文名称</div></td>
    <td width="25%" class="ButtonList"> <div align="center">中文名称</div></td>
    <td class="ButtonList"> <div align="center">英文名称</div></td>
  </tr>
<%
Dim RsClassObj,ClassSql,i
ClassSql = "Select * from FS_NewsClass"
Set RsClassObj = Conn.Execute(ClassSql)
i = 1
do while Not RsClassObj.Eof
%>
  <tr> 
    <td height="20">·<% = i & "、" %><% = RsClassObj("ClassCName") %></td>
    <td><input name="textfield" type="text" style="width:100%;" onClick="this.focus();this.select();" onMouseOver="this.focus();this.select();" onMouseOut="" value="<% = RsClassObj("ClassEName") %>" readonly></td>
<%
	RsClassObj.MoveNext
	i = i + 1
	if 	RsClassObj.Eof then Exit Do
%>
    <td>·<% = i & "、" %><% = RsClassObj("ClassCName") %></td>
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
        <div align="left"><font color="#FF0000"><strong>下载列表样式ID对应表</strong></font></div>
      </div></td>
  </tr>
  <tr> 
    <td width="50%" height="30" class="ButtonListLeft"> <div align="center">样式名称</div></td>
    <td width="25%" class="ButtonList">
<div align="center">ID</div></td>
    <td class="ButtonList">
<div align="center">样式名称</div></td>
  </tr>
<%
Dim RsDownStyleObj,DownStyleSql
DownStyleSql = "Select * from FS_DownListStyle"
Set RsDownStyleObj = Conn.Execute(DownStyleSql)
i = 1
do while Not RsDownStyleObj.Eof
%>
  <tr> 
    <td height="20">·<% = i & "、" %><% = RsDownStyleObj("Name") %></td>
    <td><input name="textfield" type="text" style="width:100%;" onClick="this.focus();this.select();" onMouseOver="this.focus();this.select();" onMouseOut="" value="<% = RsDownStyleObj("ID") %>" readonly></td>
    <td><div align="center"><span style="cursor:hand;" onClick="BrowStyle('Frame.asp?FileName=Templet_DownStyleBrow.asp&PageTitle=查看下载列表样式&ID=<% = RsDownStyleObj("ID") %>');">查看</span></div></td>
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
    <td height="30" colspan="4"> <div align="left"><strong><font color="#FF0000">商品列表样式ID对应表</font></strong></div></td>
  </tr>
  <tr> 
    <td width="51%" height="30" class="ButtonListleft"> <div align="center">样式名称</div></td>
    <td width="25%" class="ButtonList">
<div align="center">ID</div></td>
    <td class="ButtonList">
<div align="center">样式名称</div></td>
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
      ·<% = i & "、" %><% = RsMallObj("Name") %></td>
    <td bgcolor="#FFFFFF"> 
      <input name="textfield" type="text" style="width:100%;" onClick="this.focus();this.select();" onMouseOver="this.focus();this.select();" onMouseOut="" value="<% = RsMallObj("ID") %>" readonly></td>
    <td bgcolor="#FFFFFF"> 
      <div align="center"><span style="cursor:hand;" onClick="BrowStyle('Frame.asp?FileName=Templet_DownStyleBrow.asp&PageTitle=查看商品列表样式&ID=<% = RsMallObj("ID") %>');">查看</span></div></td>
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