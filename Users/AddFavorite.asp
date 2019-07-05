<% Option Explicit %>
<!--#include file="../Inc/Cls_DB.asp" -->
<!--#include file="../Inc/Const.asp" -->
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
	Dim DBC,conn,sConn
	Set DBC = new databaseclass
	Set Conn = DBC.openconnection()
	Dim I,RsConfigObj
	Set RsConfigObj = Conn.Execute("Select Domain,SiteName,UserConfer,Copyright,isEmail,isChange,IsShop from FS_Config")
	Set DBC = Nothing
%>
<!--#include file="Comm/User_Purview.Asp" -->
<%
Dim NewsId,Newsobj
NewsId = Replace(Replace(Request("NewsId"),"'",""),Chr(39),"")
Set Newsobj = Conn.execute("Select id from FS_News where id="&Replace(NewsId,"'",""))
If Newsobj.Eof then
		Response.Write("<script>alert(""找不到此新闻！"&CopyRight&""");location=""javascript:history.back()"";</script>")  
		Response.End
Else
	Dim Newssobj
	Set Newssobj = Conn.execute("Select Pid from FS_Favorite where Pid="&Replace(NewsId,"'",""))
	If Not Newssobj.eof then
		Response.Write("<script>alert(""此新闻已经在你的收藏夹中！"&CopyRight&""");location=""javascript:window.close()"";</script>")  
		Response.End
	Else
		Dim RsFObj,RsFSQL
		Set RsFObj = Server.CreateObject(G_FS_RS)
		RsFSQL = "Select * from FS_Favorite where 1=0"
		RsFObj.Open RsFSQL,Conn,1,3
		RsFObj.AddNew
		RsFObj("UserID") = Session("MemID")
		RsFObj("Pid") = Newsobj("id")
		RsFObj("isTF") = 0
		RsFObj("Addtime") = Now
		RsFObj.Update
		Set RsFObj = Nothing
		Response.Write("<script>alert(""添加到收藏成功！"&CopyRight&""");location=""User_Favorite.asp"";</script>")  
		Response.End
	End If
	Set Newssobj = Nothing
End if
%>
