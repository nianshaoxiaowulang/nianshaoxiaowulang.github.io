<% Option Explicit %>
<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Const.asp" -->
<!--#include file="../../Inc/Md5.asp" -->
<!--#include file="../../Inc/Function.asp" -->
<LINK href="../Css/UserCSS.css" type=text/css  rel=stylesheet></HEAD>
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
  Dim Email,Checkrsobj
  Email=Replace(replace(trim(Request("email")),"'","''"),Chr(39),"")
    If len(Email)<6 then 
		Response.Write("����ȷ��дemail")
		Response.end
    End if
	If Email="" then 
		Response.Write("����дemail")
		Response.end
	End If
	set Checkrsobj=conn.execute("select * from FS_Members where Email='"&Email&"'")
	If Not Checkrsobj.eof then 
		Response.Write("�Ѿ���ע��")
		Response.end
	Else
		Response.Write("��Email����ע��")
		Response.end
	End if
%>