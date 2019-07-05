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
	Set DBC = New DataBaseClass
	Set Conn = DBC.OpenConnection()
	Set DBC = Nothing
	Dim DBC,Conn,RsLoginObj,RsLogObj
	Dim MemName,Password,VerifyCode,SqlLogin,Url
	MemName = Replace(Replace(Trim(Request.Form("MemName")),"'","''"),Chr(39),"")
	Password = md5(Replace(Replace(Trim(Request.Form("Password")),"'","''"),Chr(39),""),16)
	If MemName = "" or  Password = "" then 
		Response.Write("<script>alert(""用户名和密码不能为空！"&CopyRight&""");location=""javascript:history.back()"";</script>")  
		Response.End
	End if
	Set RsLoginObj = Server.CreateObject (G_FS_RS)
	SqlLogin = "Select * From FS_Members where MemName='"&MemName&"' and  password='"&Password&"'"
	RsLoginObj.Open SqlLogin,Conn,1,1
	If Not RsLoginObj.EOF then 
		If cint(RsloginObj("Lock"))=1 then
			Response.Write("<script>alert(""您已经被管理员锁定，请与管理员联系！"&CopyRight&""");location=""javascript:history.back()"";</script>")  
			Response.End
		End if
		Session("MemName") = RsLoginObj("MemName")
		Session("sName") = RsLoginObj("Name")
		Session("MemPassword") =RsLoginObj("Password") 
		Session("MemID") = RsLoginObj("ID")
		Session("email") = RsLoginObj("Email")
		Session("GroupId") = RsLoginObj("GroupId")
		'取得Cookies
	    Response.Cookies("Foosun")("MemName") = RsLoginObj("MemName")
	    Response.Cookies("Foosun")("MemPassword") = RsLoginObj("Password") 
	    Response.Cookies("Foosun")("MemID") = RsLoginObj("ID")
	    Response.Cookies("Foosun")("email") = RsLoginObj("Email")
	    Response.Cookies("Foosun")("GroupID") = RsLoginObj("GroupID")
	    Response.Cookies("Foosun")("Point") = RsLoginObj("Point")
		Dim LoginTime
		LoginTime = date()
		Dim Rscon
		Set Rscon= conn.execute("select NumberContPoint,NumberLoginPoint from FS_Config")
		Conn.execute("update FS_members set LoginNum=LoginNum+1,Point=Point+"&clng(Rscon("NumberLoginPoint"))&",LastLoginIP='"&Request.ServerVariables("Remote_ADDR")&"',LastLoginTime='"&LoginTime&"' where MemName='"&MemName&"'")'用户登陆一次，积分+1分
		Rscon.close
		Set Rscon=nothing
		Response.Redirect("main.asp")
		Response.End
	Else
	   Response.Write("<script>alert(""非法登陆，请确认密码的正确性！"&CopyRight&""");location=""javascript:history.back()"";</script>")  
	   Response.End
	End If
	Set Conn = Nothing
	Set RsLoginObj = Nothing
%>