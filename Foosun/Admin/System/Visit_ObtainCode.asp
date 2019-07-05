<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
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
if Not JudgePopedomTF(Session("Name"),"P080501") then Call ReturnError1()
Dim TruePlusDir
If PlusDir="" then
	TruePlusDir=""
Else
	TruePlusDir="/"&PlusDir
End If

%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>流量统计代码</title>
</head>
<body topmargin="2" leftmargin="2" oncontextmenu="//return false;">
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td height="28" class="ButtonListLeft"> <div align="center"><strong>流量统计代码调用</strong></div></td>
  </tr>
</table>
<br>
<br>
<table width="85%" height="90"  border="0" align="center" cellpadding="3" cellspacing="1" bordercolor="e6e6e6" bgcolor="dddddd">
  <tr bgcolor="#FFFFFF"> 
    <td width="16%" valign="middle"> 
      <div align="center">有图标</div></td>
    <td width="84%" valign="middle"><SPAN class=small2><FONT face="Verdana, Arial, Helvetica, sans-serif">&lt;script 
      src="<%=confimsn("DoMain")%><%=TruePlusDir%>/count/count.asp?Type=Pic"&gt;&lt;/script&gt;</FONT></SPAN></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td valign="middle"> 
      <div align="center">无图标</div></td>
    <td valign="middle"><FONT face="Verdana, Arial, Helvetica, sans-serif"><SPAN class=small2>&lt;script 
      src="<%=confimsn("DoMain")%><%=TruePlusDir%>/count/count.asp"&gt;&lt;/script&gt;</SPAN></FONT></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td valign="middle"> 
      <div align="center">文字统计</div></td>
    <td valign="middle"><FONT face="Verdana, Arial, Helvetica, sans-serif"><SPAN class=small2>&lt;script 
      src="<%=confimsn("DoMain")%><%=TruePlusDir%>/count/count.asp?Type=Word"&gt;&lt;/script&gt;</SPAN></FONT></td>
  </tr>
</table>
<div align="center"></div>
</body>
</html>
