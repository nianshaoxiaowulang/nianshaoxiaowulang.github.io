<% Option Explicit %>
<!--#include file="Inc/Const.asp" -->
<!--#include file="Inc/Cls_DB.asp" -->
<!--#include file="Inc/Function.asp" -->
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
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing

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
Dim SearchYear,SearchMonth,SearchDate,RecordFileName
SearchYear = Replace(Replace(Request("SearchYear"),"'",""),Chr(39),"")
SearchMonth = Replace(Replace(Request("SearchMonth"),"'",""),Chr(39),"")
SearchDate = Replace(Replace(Request("SearchDate"),"'",""),Chr(39),"")
if SearchYear = "" then SearchYear = Year(Now)
if SearchMonth = "" then SearchMonth = Month(Now)
if SearchDate = "" then SearchDate = Day(Now)
RecordFileName = SearchYear & "-" & SearchMonth & "-" & SearchDate & ".htm"
Set Conn = Nothing
Response.Redirect(AvailableDoMain & "/" & RecordNewsListSavePath & "/" & RecordFileName)
%>