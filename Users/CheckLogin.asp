<% Option Explicit %>
<!--#include file="../Inc/Cls_DB.asp" -->
<!--#include file="../Inc/Const.asp" -->
<!--#include file="../Inc/Md5.asp" -->
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
	Set DBC = New DataBaseClass
	Set Conn = DBC.OpenConnection()
	Set DBC = Nothing
	Dim DBC,Conn,RsLoginObj,RsLogObj
	Dim MemName,Password,VerifyCode,SqlLogin,Url
	MemName = Replace(Replace(Trim(Request.Form("MemName")),"'","''"),Chr(39),"")
	Password = md5(Replace(Replace(Trim(Request.Form("Password")),"'","''"),Chr(39),""),16)
	If MemName = "" or  Password = "" then 
		Response.Write("<script>alert(""�û��������벻��Ϊ�գ�"&CopyRight&""");location=""javascript:history.back()"";</script>")  
		Response.End
	End if
	Set RsLoginObj = Server.CreateObject (G_FS_RS)
	SqlLogin = "Select * From FS_Members where MemName='"&MemName&"' and  password='"&Password&"'"
	RsLoginObj.Open SqlLogin,Conn,1,1
	If Not RsLoginObj.EOF then 
		If cint(RsloginObj("Lock"))=1 then
			Response.Write("<script>alert(""���Ѿ�������Ա�������������Ա��ϵ��"&CopyRight&""");location=""javascript:history.back()"";</script>")  
			Response.End
		End if
		Session("MemName") = RsLoginObj("MemName")
		Session("sName") = RsLoginObj("Name")
		Session("MemPassword") =RsLoginObj("Password") 
		Session("MemID") = RsLoginObj("ID")
		Session("email") = RsLoginObj("Email")
		Session("GroupId") = RsLoginObj("GroupId")
		'ȡ��Cookies
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
		Conn.execute("update FS_members set LoginNum=LoginNum+1,Point=Point+"&clng(Rscon("NumberLoginPoint"))&",LastLoginIP='"&Request.ServerVariables("Remote_ADDR")&"',LastLoginTime='"&LoginTime&"' where MemName='"&MemName&"'")'�û���½һ�Σ�����+1��
		Rscon.close
		Set Rscon=nothing
		Response.Redirect("main.asp")
		Response.End
	Else
	   Response.Write("<script>alert(""�Ƿ���½����ȷ���������ȷ�ԣ�"&CopyRight&""");location=""javascript:history.back()"";</script>")  
	   Response.End
	End If
	Set Conn = Nothing
	Set RsLoginObj = Nothing
%>