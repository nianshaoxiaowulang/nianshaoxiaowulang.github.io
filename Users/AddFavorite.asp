<% Option Explicit %>
<!--#include file="../Inc/Cls_DB.asp" -->
<!--#include file="../Inc/Const.asp" -->
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
'�������2�ο��������뾭����Ѷ��˾������������׷����������
'==============================================================================
	Dim DBC,conn,sConn
	Set DBC = new databaseclass
	Set Conn = DBC.openconnection()
	Dim I,RsConfigObj
	Set RsConfigObj = Conn.Execute("Select Domain,SiteName,UserConfer,Copyright,isEmail,isChange,IsShop from FS_Config")
	Set DBC = Nothing
%>
<!--#include file="Comm/User_Purview.Asp" -->
<%
Dim NewsId,Newsobj
NewsId = Replace(Replace(Request("NewsId"),"'",""),Chr(39),"")
Set Newsobj = Conn.execute("Select id from FS_News where id="&Replace(NewsId,"'",""))
If Newsobj.Eof then
		Response.Write("<script>alert(""�Ҳ��������ţ�"&CopyRight&""");location=""javascript:history.back()"";</script>")  
		Response.End
Else
	Dim Newssobj
	Set Newssobj = Conn.execute("Select Pid from FS_Favorite where Pid="&Replace(NewsId,"'",""))
	If Not Newssobj.eof then
		Response.Write("<script>alert(""�������Ѿ�������ղؼ��У�"&CopyRight&""");location=""javascript:window.close()"";</script>")  
		Response.End
	Else
		Dim RsFObj,RsFSQL
		Set RsFObj = Server.CreateObject(G_FS_RS)
		RsFSQL = "Select * from FS_Favorite where 1=0"
		RsFObj.Open RsFSQL,Conn,1,3
		RsFObj.AddNew
		RsFObj("UserID") = Session("MemID")
		RsFObj("Pid") = Newsobj("id")
		RsFObj("isTF") = 0
		RsFObj("Addtime") = Now
		RsFObj.Update
		Set RsFObj = Nothing
		Response.Write("<script>alert(""��ӵ��ղسɹ���"&CopyRight&""");location=""User_Favorite.asp"";</script>")  
		Response.End
	End If
	Set Newssobj = Nothing
End if
%>
