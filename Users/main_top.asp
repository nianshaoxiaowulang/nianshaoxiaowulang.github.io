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
Set DBC = Nothing
%>
<!--#include file="Comm/User_Purview.Asp" -->
<%
	Dim RsConfigObj
	Set RsConfigObj = Conn.Execute("Select SiteName,UserConfer,Copyright,isEmail,isChange,IsShop from FS_Config")
	Dim RsUserObj
	Set RsUserObj = Conn.Execute("Select Point,RegTime,UserNo,UserPoint,ShopPoint From FS_Members where MemName = '"& Session("MemName")&"' and Password = '"& Session("MemPassword") &"'")
%>
<HTML><HEAD>
<TITLE><%=RsConfigObj("SiteName")%> >> 会员中心</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="Css/UserCSS.css" type=text/css  rel=stylesheet>
</HEAD>
<body bgcolor="#F5F5F5" leftmargin="3" topmargin="0">
<fieldset>
<legend></legend>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="86%"><strong><font color="#FF0000">　</font></strong> <font color="#FF0000"><strong><%=Session("MemName")%></strong></font> 欢迎您！！<font color="#000000"> 
      <%
Dim NewsSql,GetMessageObj,TotleMessage
NewsSql = "Select * from FS_Message Where MeRead='"& session("memname")&"' and ReadTF=0 and isDelR=0 and IsRecyle=0"
Set GetMessageObj = Server.CreateObject(G_FS_RS)
GetMessageObj.Open NewsSql,Conn,1,1
TotleMessage = GetMessageObj.Recordcount
If TotleMessage=0 then
	Response.Write("<a href=User_Message.asp target=main>短消息(0)</a>")
Else
	Response.Write("<a href=User_Message.asp target=main><font color=red><b>您有新短消息("&TotleMessage&")</b></font></a>")
End If
%>
<span class="f41">，用户编号:<font color="#FF0000"><% =  RsUserObj("UserNo") %></font> 
<%
If cint(RsConfigObj("isShop"))=1 then
%>
      　可用金币:
<% =  RsUserObj("UserPoint") %>      　消费积分:<% =  RsUserObj("ShopPoint") %>
      <%End If%>
      　 注册时间: 
      <% =  RsUserObj("RegTime") %>
    </td>
    <td width="14%"><div align="center"><a href="main.asp" target="_top"><font color="#FF0000">控制面板</font></a> 
        | <a href="Comm/LetOut.asp" target="_top">退出</a></div></td>
  </tr>
</table>

</fieldset></body>
</html>
<%
Set Conn=nothing
%>
