<% Option Explicit %>
<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Const.asp" -->
<!--#include file="../../Inc/Md5.asp" -->
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
Dim  DBC,Conn
Set  DBC = New DataBaseClass
Set  Conn = DBC.OpenConnection()
Set  DBC = Nothing
Dim Pid
Pid=Replace(Replace(request("Pid"),"'",""),Chr(39),"")
Conn.execute("update FS_Shop_Products set ClickNum=ClickNum+1 where id="&Pid)
Dim Rs
Set Rs = server.createobject(G_FS_RS)
Rs.source = "select ClickNum from FS_Shop_Products where id="&pid
Rs.open rs.source,conn,1,1
If Not Rs.Eof then
%>
   javastr="<%=rs("ClickNum")%>"
   document.write(javastr)
<%
else
%>
   javastr="0"
   document.write(javastr)
<%
End if
Rs.close
set Rs=nothing
Set Conn = Nothing
%>
