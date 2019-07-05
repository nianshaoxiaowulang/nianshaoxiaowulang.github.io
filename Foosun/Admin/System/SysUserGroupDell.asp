<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
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
if Not JudgePopedomTF(Session("Name"),"P040303") then Call ReturnError()
Dim UserGroupID
UserGroupID = Request("ID")
%>
<html>
<head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>会员组删除</title>
</head>
<body leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="5" cellpadding="0">
<form action="" name="JSDellForm" method="post">
  <tr> 
    <td width="7%" height="10"></td>
    <td width="19%" rowspan="3"><div align="center"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
    <td width="68%"></td>
    <td width="6%" height="10"></td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>您确定要删除此会员组<br>并删除此会员组下的所有会员?</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td height="2"></td>
    <td height="2"></td>
    <td height="2"></td>
  </tr>
  <tr> 
    <td></td>
    <td colspan="2"><div align="center"> 
        <input type="submit" name="Submit" value=" 确 定 ">
        <input type="hidden" name="action" value="Submit">
        <input type="hidden" name="ID" value="<% = UserGroupID %>">
        <input type="button" name="Submit2" value=" 取 消 " onClick="window.close();">
      </div></td>
    <td></td>
  </tr>
  <tr> 
    <td height="10"></td>
    <td height="10" colspan="2"></td>
    <td height="10"></td>
  </tr>
</form>
</table>
</body>
</html>
<%
if Request.Form("action") = "Submit" then
	UserGroupID = Replace(UserGroupID,"***",",")
	Conn.Execute("Delete from FS_Message where MeRead in (Select MemName from FS_Members where GroupID in (Select GroupID from FS_MemGroup where ID in (" & UserGroupID & ")))")
	Conn.Execute("Delete from FS_Members where GroupID in (Select GroupID from FS_MemGroup where ID in (" & UserGroupID & "))")
	Conn.Execute("Delete from FS_MemGroup where ID in (" & UserGroupID & ")")
	Response.write("<script>dialogArguments.location.reload();window.close();</script>")
	Response.end
end if
%>