<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="inc/Config.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
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
Dim DBC,Conn,CollectConn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
DBC.ConnStr = CollectDBConnectionStr
Set CollectConn = DBC.OpenConnection()
Set DBC = Nothing
'判断权限
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P080302") then Call ReturnError()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet">
<body topmargin="0" leftmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <form action="" method="pose" name="Form">
    <tr> 
      <td width="120" height="80"> 
        <div align="center"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
      <td height="80"> 
        <div align="left">确定要删除吗？ 
          <input type="hidden" value="Submit" name="Action">
          <input type="hidden" name="NewsIDStr" value="<% = Request("NewsIDStr") %>">
        </div></td>
    </tr>
    <tr> 
      <td colspan="2"><div align="center"> 
          <input type="submit" name="Submit" value=" 确 定 ">
          <input type="button" name="Submit2" onClick="window.close();" value=" 取 消 ">
        </div></td>
    </tr>
  </form>
</table>
</body>
</html>
<%
if Request("Action") = "Submit" then
	Dim NewsIDStr,DelSql
	NewsIDStr = Request("NewsIDStr")
	if NewsIDStr <> "" then
		'On Error Resume Next
		NewsIDStr = Replace(NewsIDStr,"***",",")
		DelSql = "Delete from FS_News where ID in (" & NewsIDStr & ")"
		CollectConn.Execute(DelSql)
		if Err.Number = 0 then
			Response.Write("<script>dialogArguments.location.reload();window.close();</script>")
		else
			Response.Write("<script>dialogArguments.location.reload();window.close();</script>")
		end if
		Set CollectConn = Nothing
	end if
end if
%>