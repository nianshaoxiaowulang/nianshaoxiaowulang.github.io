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
'������ƣ���Ѷ��վ��Ϣ����ϵͳ
'��ǰ�汾��Foosun Content Manager System(FoosunCMS V3.1.0930)
'���¸��£�2005.10
'==============================================================================
'Copyright (C) 2002-2004 Foosun.Net  All rights reserved.
'��ҵע����ϵ��028-85098980-601,��Ŀ������028-85098980-606��609,�ͻ�֧�֣�608
'��Ʒ��ѯQQ��394226379,159410,125114015
'����֧��QQ��315485710,66252421 
'��Ŀ����QQ��415637671��655071
'���򿪷����Ĵ���Ѷ�Ƽ���չ���޹�˾(Foosun Inc.)
'Email:service@Foosun.cn
'MSN��skoolls@hotmail.com
'��̳֧�֣���Ѷ������̳(http://bbs.foosun.net)
'�ٷ���վ��www.Foosun.cn  ��ʾվ�㣺test.cooin.com 
'��վͨϵ��(���ܿ��ٽ�վϵ��)��www.ewebs.cn
'==============================================================================
'��Ѱ汾���ڳ�����ҳ������Ȩ��Ϣ�������ϱ�վLOGO��������
'��Ѷ��˾�����˳���ķ���׷��Ȩ��
'============================================================================================================
%>
<!--#include file="../../../Inc/Session.asp" -->
<%
'Ȩ���ж�
if Not JudgePopedomTF(Session("Name"),"P031300") then Call ReturnError1()
if Not JudgePopedomTF(Session("Name"),"P031304") then Call ReturnError1()

Dim SqlStr,StyleContent,QueryNum,ColSpan,RowSpan,ColNum,RowNum,PreviewContent
SqlStr = Request("SqlStr")
'�ж�SQL�����ĺϷ���
If InStr(Lcase(SqlStr),"select") = 0 Or InStr(Lcase(SqlStr),"insert") <> 0 Or InStr(Lcase(SqlStr),"drop") <> 0 Or InStr(Lcase(SqlStr),"update") <> 0 Then
	Response.Write("�Ƿ�Sql���")
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
			 ��ѯ����:<input type="text" name = "QueryNum" value="<%=QueryNum%>" style="width:10%">
			 ˮƽ���:<input type="text" name = "ColSpan" value="<%=ColSpan%>" style="width:10%">
			 ��ֱ���:<input type="text" name = "RowSpan" value="<%=RowSpan%>" style="width:10%">
			 ��&nbsp;&nbsp;&nbsp;&nbsp;��:<input type="text" name = "RowNum" value="<%=RowNum%>" style="width:10%">
			 ��&nbsp;&nbsp;&nbsp;&nbsp;��:<input type="text" name = "ColNum" value="<%=ColNum%>" style="width:10%">
			 <input type="submit" value="ˢ��">
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