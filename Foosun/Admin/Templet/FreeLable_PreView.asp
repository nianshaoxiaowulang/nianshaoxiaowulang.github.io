<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/checkPopedom.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../Refresh/Function.asp" -->
<!--#include file="../Refresh/RefreshFunction.asp" -->
<% 
Dim  DBC,Conn,TempClassListStr,TempListStr
Set  DBC = New DataBaseClass
Set  Conn = DBC.OpenConnection()
Set  DBC = Nothing
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
'============================================================================================================
%>
<!--#include file="../../../Inc/Session.asp" -->
<%
'权限判断
if Not JudgePopedomTF(Session("Name"),"P031300") then Call ReturnError1()
if Not JudgePopedomTF(Session("Name"),"P031304") then Call ReturnError1()

Dim SqlStr,StyleContent,QueryNum,ColSpan,RowSpan,ColNum,RowNum,PreviewContent
SqlStr = Request("SqlStr")
'判断SQL语名的合法性
If InStr(Lcase(SqlStr),"select") = 0 Or InStr(Lcase(SqlStr),"insert") <> 0 Or InStr(Lcase(SqlStr),"drop") <> 0 Or InStr(Lcase(SqlStr),"update") <> 0 Then
	Response.Write("非法Sql语句")
	Response.End
End if
'response.Write(SqlStr)
'response.End()
StyleContent = Replace(Replace(Replace(Request("StyleContent"),"%3C","<"),"%3E",">"),"%3D","""")
QueryNum = Request("QueryNum")
ColSpan = Request("ColSpan")
RowSpan = Request("RowSpan")
ColNum = Request("ColNum")
RowNum = Request("RowNum")
If Not IsNumeric(QueryNum) Then 
	QueryNum = 4
Elseif Cint(QueryNum) <= 0 Then
	QueryNum = 4
Elseif Cint(QueryNum) > 100 Then
	QueryNum = 100
End if

If Not IsNumeric(ColNum) Then 
	ColNum = 1
Elseif Cint(ColNum) <= 0 Then
	ColNum = 1
Elseif Cint(ColNum) > 4 Then
	ColNum = 4
End if

If Not IsNumeric(RowNum) Then 
	RowNum = 1
Elseif Cint(RowNum) <= 0 Then
	RowNum = 1
Elseif Cint(RowNum) > 4 Then
	RowNum = 4
End if

If Not IsNumeric(ColSpan) Then 
	ColSpan = ""
Elseif Cint(ColSpan) <= 0 Then
	ColSpan = ""
End if
If Not IsNumeric(RowSpan) Then 
	RowSpan = ""
Elseif Cint(RowSpan) <= 0 Then
	RowSpan = ""
End if

GetAvailableDoMain
PreviewContent = CreateFreeLable(SqlStr,StyleContent,QueryNum,"","",ColSpan,RowSpan,ColNum,RowNum)

Set conn = nothing
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<link href="../../../CSS/Style.css" rel="stylesheet">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body leftmargin="0" topmargin="0">
<form action="" mothed="post" name="PreViewForm">
<div align=center valign=middle>
<table width=90% height=90% bgcolor="#cccccc">
	<tr bgcolor="#ffffff">
		<td align="center" valign="middle" height=20>
			 查询数量:<input type="text" name = "QueryNum" value="<%=QueryNum%>" style="width:10%">
			 水平间距:<input type="text" name = "ColSpan" value="<%=ColSpan%>" style="width:10%">
			 垂直间距:<input type="text" name = "RowSpan" value="<%=RowSpan%>" style="width:10%">
			 行&nbsp;&nbsp;&nbsp;&nbsp;数:<input type="text" name = "RowNum" value="<%=RowNum%>" style="width:10%">
			 列&nbsp;&nbsp;&nbsp;&nbsp;数:<input type="text" name = "ColNum" value="<%=ColNum%>" style="width:10%">
			 <input type="submit" value="刷新">
			 <input type="hidden" name="SqlStr" value = "<%=SqlStr%>">
			 <input type="hidden" name="StyleContent" value = "<%=Replace(Replace(Replace(StyleContent,"<","%3C"),">","%3E"),"""","%3D")%>">
		</td>
	</tr>
	<tr bgcolor="#ffffff">
		<td align="center" valign="middle">
			<%=PreviewContent%>
		</td>
	</tr>
</table>
</div>
</form>
</body>
</html>
<script>
</script>