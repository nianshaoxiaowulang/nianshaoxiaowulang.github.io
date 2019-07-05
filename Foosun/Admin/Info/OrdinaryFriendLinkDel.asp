<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
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

%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P070503") then Call ReturnError()
Dim FriendLinkID
FriendLinkID = Request("FriendLinkID")
%>
<html>
<head>
<link rel="stylesheet" href="../../../CSS/ModeWindow.css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>友情链接管理</title>
</head>
<body topmargin="0" leftmargin="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="3">
  <form method="post" name="DelForm">
    <tr> 
      <td><div align="center"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
      <td colspan="2">您确定要删除此友情链接?</td>
    </tr>
    <tr> 
      <td colspan="3"><div align="center"> 
          <input type="submit" name="Submit" value=" 确 定 ">
          <input type="hidden" name="Action" value="Submit">
          <input type="hidden" name="FriendLinkID" value="<% = FriendLinkID %>">
          <input type="button" onClick="window.close();" name="Submit2" value=" 取 消 " >
      </div></td>
    </tr>
  </form>
</table>
</body>
</html>
<%
if Request.Form("Action") = "Submit" then
	Conn.Execute("Delete from FS_FriendLink where ID in (" & Replace(FriendLinkID,"***",",") & ")")
	Response.Write("<script>dialogArguments.location.reload();window.close();</script>")
end if
Set Conn = Nothing
%>
