<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../Refresh/Function.asp" -->
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
if Not JudgePopedomTF(Session("Name"),"P010000") then Call ReturnError()
Dim AvailableDoMain
GetAvailableDoMain
Sub GetAvailableDoMain()
	Dim ConfigSql,RsConfigObj
	ConfigSql = "Select DoMain,MakeType,IndexExtName from FS_Config"
	Set RsConfigObj = Conn.Execute(ConfigSql)
	if Not RsConfigObj.Eof then
		AvailableDoMain = RsConfigObj("DoMain")
	else
		AvailableDoMain = GetDoMain
	end if
	Set RsConfigObj = Nothing
End Sub

Dim ID,Table,Sql,ReadConfigObj,RsReadObj
ID = Request("ID")
Table = Request("Table")
If ID = "" then
	Response.Write("<script>alert(""参数错误"");window.close();</script>")
	Response.end
end if
if Table = "News" then
	if Not JudgePopedomTF(Session("Name"),"P010511") then Call ReturnError()
	Sql = "Select * from FS_News  where NewsID='" & ID & "'"
elseif Table = "DownLoad" then
	if Not JudgePopedomTF(Session("Name"),"P010705") then Call ReturnError()
	Sql = "Select * from FS_DownLoad  where DownLoadID='" & ID & "'"
elseif Table = "Product" then
	if Not JudgePopedomTF(Session("Name"),"P010805") then Call ReturnError()
	sql="Select * from FS_Shop_Products where id="&ID
elseif Table = "NewsClass" then
	sql="Select * from FS_NewsClass where ClassID='"&ID&"'"
else
	Response.Write("<script>alert(""参数错误"");window.close();</script>")
	Response.end
end if
Set ReadConfigObj = Conn.Execute("Select DoMain from FS_Config")
Set RsReadObj = Server.CreateObject(G_FS_RS)
RsReadObj.Open Sql,Conn,3,3
if RsReadObj.Eof then
	Set ReadConfigObj = Nothing
	Set RsReadObj = Nothing
	Set Conn = Nothing
	Response.Write("<script>alert(""参数错误"");window.close();</script>")
	Response.end
Else
	Dim RsClassObj,URL
	Sql = "Select * from FS_NewsClass where ClassID='" & RsReadObj("ClassID") & "'"
	Set RsClassObj = Server.CreateObject(G_FS_RS)
	RsClassObj.Open Sql,Conn,1,1
	if RsClassObj.Eof then
		Set RsClassObj = Nothing
		Set ReadConfigObj = Nothing
		Set RsReadObj = Nothing
		Set Conn = Nothing
		Response.Write("<script>alert(""参数错误"");window.close();</script>")
		Response.end
	else
		if Not JudgePopedomTF(Session("Name"),"" & RsClassObj("ClassID") & "") then Call ReturnError1()
		if Table = "News" then
			URL = GetOneNewsLinkURL(ID)
		elseif Table = "DownLoad" then
			URL = GetOneDownLoadLinkURL(ID)
		elseif Table="Product" then 
			URl=GetOneProductLinkURL(ID)
		elseif Table="NewsClass" then 
			URl=GetOneClassLinkURLByID(ID)
		end If
		if URL = "" then
			Response.Write("没有审核，还不能够预览......")
		else
			Response.Redirect(URL)
		end if
	end if
end if
Set RsClassObj = Nothing
Set ReadConfigObj = Nothing
Set RsReadObj = Nothing
Set Conn = Nothing
%>
