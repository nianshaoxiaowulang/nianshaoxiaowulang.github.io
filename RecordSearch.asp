<% Option Explicit %>
<!--#include file="Inc/Const.asp" -->
<!--#include file="Inc/Cls_DB.asp" -->
<!--#include file="Inc/Function.asp" -->
<%
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