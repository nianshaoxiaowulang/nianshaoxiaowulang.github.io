<% Option Explicit %>
<!--#include file="../Inc/Function.asp" -->
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
	Dim DBC,conn
	Set DBC = new databaseclass
	Set conn = DBC.openconnection()
	Set DBC = nothing
	
	Dim I,RsConfigObj,VerifyCode
	Set RsConfigObj = Conn.Execute("Select SiteName,UserConfer,Copyright,isEmail,SendPoint from FS_Config")
	VerifyCode = Replace(Trim(Request("Ver")),"'","")
	'-------------------判断用户提交的资料。并写入数据库
	If  VerifyCode <> CStr(Session("GetCode"))  then 
		Response.Write("<script>alert(""错误：\n验证码错误"");location.href=""javascript:history.back()"";</script>")
		Response.End
	End if
	If VerifyCode = "" then
		Response.Write("<script>alert(""错误：\n请填写验证码"");location.href=""javascript:history.back()"";</script>")
		Response.End
	Elseif Session("GetCode") = "9999" then
		Session("GetCode")=""
	Elseif Session("GetCode") = "" then
		Response.Write("<script>alert(""错误：\n请不要重复提交，如需重新登录请返回登录页面。\n返回后请刷新登录页面后重新输入正确的信息"");location.href=""javascript:history.back()"";</script>")
		Response.End
	Elseif cstr(Session("GetCode"))<>cstr(Trim(Request("ver"))) then
		Response.Write("<script>alert(""错误：\n您输入的确认码和系统产生的不一致，请重新输入。\n返回后请刷新登录页面后重新输入正确的信息"");location.href=""javascript:history.back()"";</script>")
		Response.End
	End if
	Session("GetCode")=""
	Dim RsMemberObj,RsMemberObj1
	Set RsMemberObj = Conn.Execute("Select MemName from FS_members where MemName = '" & replace(request.Form("Username"),"'","") &"'")    
	If Not RsMemberObj.Eof then
		Response.Write("<script>alert(""用户名已经存在！请重新选择"&CopyRight&""");location=""javascript:history.back()"";</script>")
		Response.End
	End if
	Set RsMemberObj1 = Conn.Execute("Select Email from FS_members where Email ='" & trim(replace(request.Form("email"),"'","")) &"'")
	If Not RsMemberObj1.Eof then
		Response.Write("<script>alert(""电子邮件已经存在！请重新选择"&CopyRight&""");location=""javascript:history.back()"";</script>")
		Response.End
	End if
	'response.write request.form
	'response.end
	Randomize 
	Dim RandomFigure
	RandomFigure = CStr(Int((9999 * Rnd) + 1))
	Dim RsUserAddObj,RsUserSql
	Set RsUserAddObj = Server.CreateObject(G_FS_RS)
	RsUserSql = "Select * from FS_Members where 1=0"
	RsUserAddObj.Open RsUserSql,Conn,3,3
	RsUserAddObj.AddNew
	RsUserAddObj("MemName") = NoCSSHackInput(Replace(Replace(Request.Form("Username"),"""",""),"'",""))
	RsUserAddObj("Password") = md5(Request.Form("sPassword"),16)
	RsUserAddObj("Email") = NoCSSHackInput(Replace(Replace(Request.Form("email"),"""",""),"'",""))
	If Replace(Replace(Request.Form("tel"),"""",""),"'","")<>"" then
		RsUserAddObj("Telephone") = NoCSSHackInput(Replace(Replace(Request.Form("tel"),"""",""),"'",""))
	End If
	If Replace(Replace(Request.Form("sName"),"""",""),"'","")<>"" then
		RsUserAddObj("Name") = NoCSSHackInput(Replace(Replace(Request.Form("sName"),"""",""),"'",""))
	End If
	If Request.Form("Sex")="1" then
		RsUserAddObj("Sex") = 1
	Else
		RsUserAddObj("Sex") = 0
	End If
	If Replace(Replace(Request.Form("Address"),"""",""),"'","")<>"" then
		RsUserAddObj("Address") = NoCSSHackInput(Replace(Replace(Request.Form("Address"),"""",""),"'",""))
	End If
	Dim iPostCode
	iPostCode=NoCSSHackInput(Replace(Replace(Request.Form("PostCode"),"""",""),"'",""))
	If iPostCode="" or IsNull(iPostCode) then iPostCode=0
	RsUserAddObj("PostCode") = iPostCode
	RsUserAddObj("Birthday") = NoCSSHackInput(Request.Form("yyear")&"-"&Request.Form("mmonth")&"-"&Request.Form("dday"))
	RsUserAddObj("Province") = NoCSSHackInput(Replace(Replace(Request.Form("Province"),"""",""),"'",""))
	If Request.Form("City")<>"" then
		RsUserAddObj("City") = NoCSSHackInput(Replace(Replace(Request.Form("City"),"""",""),"'",""))
	End if
	RsUserAddObj("PassQuestion") = NoCSSHackInput(Replace(Replace(Request.Form("PassQuestion"),"""",""),"'",""))
	RsUserAddObj("PassAnswer") = md5(Request.Form("PassAnswer"),16)
	RsUserAddObj("VerGetType") = NoCSSHackInput(Replace(Replace(Request.Form("VerGetType"),"""",""),"'",""))
	if Request.Form("VerGetCode")<>"" then
		RsUserAddObj("VerGetCode") = md5(Replace(Replace(Request.Form("VerGetCode"),"""",""),"'",""),32)
	End if
	RsUserAddObj("RegTime") = Now()
	RsUserAddObj("LastLoginIP") = Request.ServerVariables("Remote_ADDR")
	RsUserAddObj("LastLoginTime") = Now()
	RsUserAddObj("LoginNum") = 1
	RsUserAddObj("GroupID") = 0
	RsUserAddObj("Point") = 1        '积分
	RsUserAddObj("UserNo") = year(now)&month(now)&day(now)&hour(now)&RandomFigure '会员编号
	RsUserAddObj("UserPoint") = 0 
	RsUserAddObj("ShopPoint") =RsConfigObj("SendPoint") '注册送金币
	RsUserAddObj.Update
	'---
	Dim GetMemberSessionObj
	Set GetMemberSessionObj = Conn.execute("Select * from FS_Members where MemName='"& NoCSSHackInput(Replace(Replace(Request.Form("Username"),"""",""),"'","")) &"' and Password='"& md5(Request.Form("sPassword"),16) &"'")
	If Not GetMemberSessionObj.Eof then
		Session("MemName") = GetMemberSessionObj("MemName")
		Session("sName") = GetMemberSessionObj("Name")
		Session("MemPassword") =GetMemberSessionObj("Password") 
		Session("VerPassword") = Request.Form("sPassword")
		Session("MemID") = GetMemberSessionObj("ID")
		Session("email") = GetMemberSessionObj("Email")
		'取得Cookies
	    Response.Cookies("Foosun")("MemName") = GetMemberSessionObj("MemName")
	    Response.Cookies("Foosun")("MemPassword") = GetMemberSessionObj("Password") 
	    Response.Cookies("Foosun")("MemID") = GetMemberSessionObj("ID")
	    Response.Cookies("Foosun")("email") = GetMemberSessionObj("Email")
	    Response.Cookies("Foosun")("GroupID") = GetMemberSessionObj("GroupID")
	    Response.Cookies("Foosun")("Point") = GetMemberSessionObj("Point")
		GetMemberSessionObj.Close
		Set GetMemberSessionObj = Nothing
		Response.Redirect("Register_Success.Asp")
		Response.End
	Else
		Response.Write("Error(0,10000Reg)")
		Response.end
	End if
		Set RsUserAddObj = Nothing
		Set RsConfigObj = Nothing
%>