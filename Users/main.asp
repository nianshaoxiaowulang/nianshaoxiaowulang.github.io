<% Option Explicit %>
<!--#include file="../Inc/Cls_DB.asp" -->
<!--#include file="../Inc/Const.asp" -->
<!--#include file="../Inc/Md5.asp" -->
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
'==============================================================================
Dim DBC,conn,sConn
Set DBC = new databaseclass
Set Conn = DBC.openconnection()
Dim I,RsConfigObj
Set RsConfigObj = Conn.Execute("Select SiteName,Copyright,IsShop from FS_Config")
Set DBC = Nothing
%>
<!--#include file="Comm/User_Purview.Asp" -->
<html>
<head>
<title><%=RsConfigObj("SiteName")%> >>User Manage Center</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="Foosun,CMS,Foosun Content Manager System,风讯,风讯网站管理系统,风讯站点管理系统,风讯新闻系统,风讯图文系统,ASP,SQL Server,风讯商业版本,风讯图片系统,风讯科技,风讯公司,新闻,论坛,BBS,讨论,文章系统,风讯加油站,风讯开发者网络,风讯技术,技术文章,数据库技术,">
</head>
<frameset rows="38,*,1,0" border="0" framespacing="0" frameborder="0">
	<frame src="main_Top.asp" name="top" scrolling="NO" noresize>		
	<frameset cols="228,*,0,0,0,0" border="0" framespacing="0" frameborder="0" id="lkoamenu_frame" name="lkoamenu_frame"> 
	<frame src="main_left.asp" name="left" noresize onchange="check()">
	<frame src="main_main.asp" name="main">	
	<frame src="bottom.asp" frameborder="NO" scrolling="NO" name="bottom" marginwidth="0" marginheight="0" style="BORDER-right: #aeaeae 1px solid;BORDER-BOTTOM: #aeaeae 1px solid;BORDER-left: #aeaeae 1px solid;">
	

</html>
<%
RsConfigObj.Close
Set RsConfigObj = Nothing
Set Conn=nothing
%>