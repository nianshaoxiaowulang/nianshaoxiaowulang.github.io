<% Option Explicit %>
<!--#include file="../Inc/Function.asp" -->
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
	Dim DBC,conn
	Set DBC = new databaseclass
	Set conn = DBC.openconnection()
	Set DBC = nothing
	
	Dim I,RsConfigObj,VerifyCode
	Set RsConfigObj = Conn.Execute("Select SiteName,UserConfer,Copyright,isEmail,SendPoint from FS_Config")
	VerifyCode = Replace(Trim(Request("Ver")),"'","")
	'-------------------�ж��û��ύ�����ϡ���д�����ݿ�
	If  VerifyCode <> CStr(Session("GetCode"))  then 
		Response.Write("<script>alert(""����\n��֤�����"");location.href=""javascript:history.back()"";</script>")
		Response.End
	End if
	If VerifyCode = "" then
		Response.Write("<script>alert(""����\n����д��֤��"");location.href=""javascript:history.back()"";</script>")
		Response.End
	Elseif Session("GetCode") = "9999" then
		Session("GetCode")=""
	Elseif Session("GetCode") = "" then
		Response.Write("<script>alert(""����\n�벻Ҫ�ظ��ύ���������µ�¼�뷵�ص�¼ҳ�档\n���غ���ˢ�µ�¼ҳ�������������ȷ����Ϣ"");location.href=""javascript:history.back()"";</script>")
		Response.End
	Elseif cstr(Session("GetCode"))<>cstr(Trim(Request("ver"))) then
		Response.Write("<script>alert(""����\n�������ȷ�����ϵͳ�����Ĳ�һ�£����������롣\n���غ���ˢ�µ�¼ҳ�������������ȷ����Ϣ"");location.href=""javascript:history.back()"";</script>")
		Response.End
	End if
	Session("GetCode")=""
	Dim RsMemberObj,RsMemberObj1
	Set RsMemberObj = Conn.Execute("Select MemName from FS_members where MemName = '" & replace(request.Form("Username"),"'","") &"'")    
	If Not RsMemberObj.Eof then
		Response.Write("<script>alert(""�û����Ѿ����ڣ�������ѡ��"&CopyRight&""");location=""javascript:history.back()"";</script>")
		Response.End
	End if
	Set RsMemberObj1 = Conn.Execute("Select Email from FS_members where Email ='" & trim(replace(request.Form("email"),"'","")) &"'")
	If Not RsMemberObj1.Eof then
		Response.Write("<script>alert(""�����ʼ��Ѿ����ڣ�������ѡ��"&CopyRight&""");location=""javascript:history.back()"";</script>")
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
	RsUserAddObj("Point") = 1        '����
	RsUserAddObj("UserNo") = year(now)&month(now)&day(now)&hour(now)&RandomFigure '��Ա���
	RsUserAddObj("UserPoint") = 0 
	RsUserAddObj("ShopPoint") =RsConfigObj("SendPoint") 'ע���ͽ��
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
		'ȡ��Cookies
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