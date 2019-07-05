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
Dim Sql,Result,ExeResult,ExeResultNum,ExeSelectTF,ErrorTF,FiledObj
Dim I,ErrObj
Result = Request.Form("Result")
if Result = "Submit" then
	Set DBC = New DataBaseClass
	Set Conn = DBC.OpenConnection()
	Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P040605") then Call ReturnError1()
	Sql = Request.Form("Sql")
	if (Sql <> "") then
		If Instr(1,lcase(Sql),"delete from FS_log")<>0 then
			Sql="Select top 10 * from FS_log order by id desc"
		End If
		ExeSelectTF = (LCase(Left(Trim(Sql),6)) = "select")
		Conn.Errors.Clear
		On Error Resume Next
		if ExeSelectTF = True then
			Set ExeResult = Conn.ExeCute(Sql,ExeResultNum)
		else
			Conn.ExeCute Sql,ExeResultNum
		end if
		If Conn.Errors.Count<>0 Then
			ErrorTF = True
			Set ExeResult = Conn.Errors
		Else
			ErrorTF = False
		End If
	end if
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>执行结果</title>
</head>
<style type="text/css">
<!--
.SysParaButtonStyle {
	border-top-width: 1px;
	border-right-width: 1px;
	border-bottom-width: 1px;
	border-left-width: 1px;
	border-right-style: solid;
	border-bottom-style: solid;
	border-left-style: solid;
	border-right-color: #999999;
	border-bottom-color: #999999;
	border-left-color: #FFFFFF;
	background-color: #E6E6E6;
}
-->
</style>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<body topmargin="2" leftmargin="2" onselectstart="//return false;" oncontextmenu="return false;">
<%
if Result = "Submit" then
if ErrorTF = True then
%>
<table width="100%" cellpadding="0" cellspacing="1" bgcolor="#000000">
  <tr> 
    <td height="20" nowrap class="SysParaButtonStyle"> 
      <div align="center">错误号</div></td>
    <td height="20" nowrap class="SysParaButtonStyle"> 
      <div align="center">来源</div></td>
    <td height="20" nowrap class="SysParaButtonStyle"> 
      <div align="center">描述</div></td>
    <td height="20" nowrap class="SysParaButtonStyle"> 
      <div align="center">帮助</div></td>
    <td height="20" nowrap class="SysParaButtonStyle"> 
      <div align="center">帮助文档</div></td>
  </tr>
  <%
	For I=1 To Conn.Errors.Count
		Set ErrObj=Conn.Errors(I-1)
%>
  <tr bgcolor="#FFFFFF"> 
    <td nowrap> 
      <% = ErrObj.Number %> </td>
    <td nowrap> 
      <% = ErrObj.Description %> </td>
    <td nowrap> 
      <% = ErrObj.Source %> </td>
    <td nowrap> 
      <% = ErrObj.Helpcontext %> </td>
    <td nowrap> 
      <% = ErrObj.HelpFile %> </td>
  </tr>
  <%
	next
%>
</table>
<%
else
%>
<table border="0" cellpadding="0" cellspacing="1" bgcolor="#000000">
  <%
	if ExeSelectTF = True then
%>
  <tr>
<%
		For Each FiledObj In ExeResult.Fields
%>
    <td nowrap class="ButtonListLeft" height="26"><div align="center">
        <% = FiledObj.name %>
      </div></td>
<%
		next
%>
  </tr>
<%
		do while Not ExeResult.Eof
%>
  <tr>
<%
			For Each FiledObj In ExeResult.Fields
%>
    <td nowrap bgcolor="#FFFFFF"> 
      <div align="center">
        <%
		 if IsNull(FiledObj.value) then
		 	Response.Write("&nbsp;")
		 else
		 	Response.Write(FiledObj.value)
		 end if
		 %>
      </div></td>
<%
			next
%>
  </tr>
<%
			ExeResult.MoveNext
		loop
	else
%>
  <tr>
    <td class="ButtonListLeft" height="26">
<div align="center">执行结果</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">
<div align="center">
        <% = ExeResultNum & "条纪录被影响"%>
      </div></td>
  </tr>
<%
	end if
%>
</table>
<%
end if
end if
%>
<form name="ExecuteForm" method="post" action="">
  <input type="hidden" name="Sql">
  <input type="hidden" name="Result">
</form>
</body>
</html>
<%
Set Conn = Nothing
Set ExeResult = Nothing
%>
