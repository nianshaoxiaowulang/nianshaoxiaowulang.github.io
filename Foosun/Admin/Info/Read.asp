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
'�������2�ο��������뾭����Ѷ��˾������������׷����������
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
	Response.Write("<script>alert(""��������"");window.close();</script>")
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
	Response.Write("<script>alert(""��������"");window.close();</script>")
	Response.end
end if
Set ReadConfigObj = Conn.Execute("Select DoMain from FS_Config")
Set RsReadObj = Server.CreateObject(G_FS_RS)
RsReadObj.Open Sql,Conn,3,3
if RsReadObj.Eof then
	Set ReadConfigObj = Nothing
	Set RsReadObj = Nothing
	Set Conn = Nothing
	Response.Write("<script>alert(""��������"");window.close();</script>")
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
		Response.Write("<script>alert(""��������"");window.close();</script>")
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
			Response.Write("û����ˣ������ܹ�Ԥ��......")
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
