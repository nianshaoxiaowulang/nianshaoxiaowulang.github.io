<% Option Explicit %>
<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Const.asp" -->
<!--#include file="../../Inc/Md5.asp" -->
<!--#include file="../../Inc/Function.asp" -->
<LINK href="../Css/UserCSS.css" type=text/css  rel=stylesheet></HEAD>
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
 Dim DBC,conn
  Set DBC = new databaseclass
  Set conn = DBC.openconnection()
  Set DBC = nothing
  Dim Username
  Username=Replace(replace(trim(Request("Username")),"'","''"),Chr(39),"")
    If len(Username)<3 then 
		Response.Write("用户名不能少于3位")
		Response.end
    End if
	If Username="" then 
		Response.Write("请填写用户名")
		Response.end
	End if
	Dim checkrsobj
	Set checkrsobj=conn.execute("select * from FS_Members where MemName='"&Username&"'")
	If Not checkrsobj.eof then 
		Response.Write("已经被注册")
		Response.end
	Else
		Response.Write("此用户名可以注册")
		Response.end
	End if
%>