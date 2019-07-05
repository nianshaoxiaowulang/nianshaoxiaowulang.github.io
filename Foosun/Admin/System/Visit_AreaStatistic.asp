<% Option Explicit %>
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
if Not JudgePopedomTF(Session("Name"),"P080508") then Call ReturnError1()
%>
<html>
<head>
<title>访问者地区统计</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script src="../JS/PublicJS.js" language="JavaScript"></script>
<body bgcolor="#FFFFFF" topmargin="2" leftmargin="2">
<table height="26" width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
      <td height="28" class="ButtonListLeft"><div align="center"><strong>访问者地区统计</strong></div></td>
  </tr>
</table>
<%
Dim RsAreaObj,Sql
Set RsAreaObj = Server.CreateObject(G_FS_RS)
Sql="Select Area From FS_FlowStatistic"
RsAreaObj.Open Sql,Conn,3,3
Dim AreaType
Dim NumIn,NumOut,NumOther
NumIn=0
NumOut=0
NumOther=0
Do While not RsAreaObj.Eof
	AreaType= RsAreaObj("Area")
	Select Case AreaType
	Case "局域网内部网"
		NumIn=NumIn+1
	Case "未知区域"
		NumOther=NumOther+1
	Case Else
		NumOut=NumOut+1
	End Select
	RsAreaObj.MoveNext
Loop
%>
<%
Dim AllNum
AllNum=NumIn+NumOut+NumOther
%>
<table width=96% border=0 cellpadding=2>
	<tr>
		<td align=center>访问者地区统计图表</td>
	</tr>
	<tr>
		<td align=center>
			<table align=center>
        <tr valign=bottom >
					
          <td nowap>内部网</td>
					
          <td align=left background=../../Images/Visit/tu_back_2.gif valign=middle width=150><img src=../../Images/Visit/bar2.gif height=15></td>
					<td nowap><% =NumIn %></td>
				</tr>
				<tr valign=bottom >
					
          <td nowap>外部网</td>
					
          <td align=left background=../../Images/Visit/tu_back_2.gif valign=middle width=150><img src=../../Images/Visit/bar2.gif height=15></td>
					<td nowap><% =NumOut %></td>
				</tr>
				<tr valign=bottom >
					
          <td align="right" nowap>未知</td>
					
          <td align=left background=../../Images/Visit/tu_back_2.gif valign=middle width=150><img src=../../Images/Visit/bar2.gif height=15></td>
					<td nowap><% =NumOther %></td>
				</tr>
				<tr valign=cente>
					<td align=right nowap>共</td>
					
          <td align=left background=../../Images/Visit/tu_back_2.gif valign=middle width=150><img src=../../Images/Visit/bar2.gif width="150" height=15></td>
					<td nowap><% = AllNum %></td>
			</table><br>
		</td>
	</tr>
</table>