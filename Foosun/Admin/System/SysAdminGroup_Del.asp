<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp"-->
<%
'==============================================================================
'软件名称：风讯网站信息管理系统
'当前版本：Foosun Content Manager System(FoosunCMS V3.1.0930)
'最新更新：2005.10
'==============================================================================
'Copyright (C) 2002-2004 Foosun.Net  All rights reserved.
'商业注册联系：028-85098980-601,项目开发：028-85098980-606、609,客户支持：608
'产品咨询QQ：394226379,159410,125114015
'技术支持QQ：315485710,66252421 
'项目开发QQ：415637671，655071
'程序开发：四川风讯科技发展有限公司(Foosun Inc.)
'Email:service@Foosun.cn
'MSN：skoolls@hotmail.com
'论坛支持：风讯在线论坛(http://bbs.foosun.net)
'官方网站：www.Foosun.cn  演示站点：test.cooin.com 
'网站通系列(智能快速建站系列)：www.ewebs.cn
'==============================================================================
'免费版本请在程序首页保留版权信息，并做上本站LOGO友情连接
'风讯公司保留此程序的法律追究权利
'如需进行2次开发，必须经过风讯公司书面允许。否则将追究法律责任
'==============================================================================
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->

<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P040103") then Call ReturnError()
Dim Result,ID
ID = Request("ID")
Result = Request.Form("Result")
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>管理员组删除</title>
</head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet">
<body topmargin="0" leftmargin="0">
<table width="96%" border="0" align="center">
  <form name="" action="" method="post">
  <tr> 
    <td width="4%" align="center" valign="middle">&nbsp;</td>
      <td width="24%" rowspan="3" align="center" valign="middle"><img src="../../Images/Question.gif" width="39" height="37"></td>
    <td width="72%" align="left" valign="middle">&nbsp;</td>
  </tr>
  <tr>
    <td align="center" valign="middle">&nbsp;</td>
    <td width="72%" align="left" valign="middle">您确定要删除此管理员组<br>
并删除此管理员组下的所有管理员?</td>
  </tr>
  <tr>
    <td align="center" valign="middle">&nbsp;</td>
    <td width="72%" align="left" valign="middle">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="3"><div align="center"> 
          <input name="Submit" type="submit" id="Submit" value=" 确 定 ">
        <input name="Result" type="hidden" id="Result" value="Submit">
          <input name="Submit1"  onClick="window.close();"type="reset" id="Submit1" value=" 取 消 ">
      </div></td>
  </tr></form>
</table>
</body>
</html>
<%
if Result= "Submit" then
	Dim Sql
	if ID <> "" then
		ID = Replace(ID,"***",",")
		Sql = "Delete from FS_Admin Where GroupID in (" & ID & ")"
		Conn.Execute(Sql)
		Sql = "Delete from FS_AdminGroup Where ID in (" & ID & ")"
		Conn.Execute(Sql)
		Set Conn = Nothing
	end if
	if Err.Number = 0 then
		%>
		<script language="JavaScript">
			dialogArguments.location.reload();
			window.close();
		</script>
		<%
	else
		%>
		<script language="JavaScript">
			alert('删除失败');
			dialogArguments.location.reload();
			window.close();
		</script>
		<%
	end if
end if
Set Conn = Nothing
%>