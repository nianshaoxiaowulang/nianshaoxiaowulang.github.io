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
Dim DBC,conn,sConn
Set DBC = new databaseclass
Set Conn = DBC.openconnection()
Dim I,RsConfigObj
Set RsConfigObj = Conn.Execute("Select SiteName,Copyright,IsShop from FS_Config")
Set DBC = Nothing
%>
<!--#include file="Comm/User_Purview.Asp" -->
<html>
<head>
<title><%=RsConfigObj("SiteName")%> >>User Manage Center</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="Foosun,CMS,Foosun Content Manager System,��Ѷ,��Ѷ��վ����ϵͳ,��Ѷվ�����ϵͳ,��Ѷ����ϵͳ,��Ѷͼ��ϵͳ,ASP,SQL Server,��Ѷ��ҵ�汾,��ѶͼƬϵͳ,��Ѷ�Ƽ�,��Ѷ��˾,����,��̳,BBS,����,����ϵͳ,��Ѷ����վ,��Ѷ����������,��Ѷ����,��������,���ݿ⼼��,">
</head>
<frameset rows="38,*,1,0" border="0" framespacing="0" frameborder="0">
	<frame src="main_Top.asp" name="top" scrolling="NO" noresize>		
	<frameset cols="228,*,0,0,0,0" border="0" framespacing="0" frameborder="0" id="lkoamenu_frame" name="lkoamenu_frame"> 
	<frame src="main_left.asp" name="left" noresize onchange="check()">
	<frame src="main_main.asp" name="main">	
	<frame src="bottom.asp" frameborder="NO" scrolling="NO" name="bottom" marginwidth="0" marginheight="0" style="BORDER-right: #aeaeae 1px solid;BORDER-BOTTOM: #aeaeae 1px solid;BORDER-left: #aeaeae 1px solid;">
	

</html>
<%
RsConfigObj.Close
Set RsConfigObj = Nothing
Set Conn=nothing
%>