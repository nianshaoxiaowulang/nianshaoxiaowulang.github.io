<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../../../Inc/NosqlHack.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
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
Dim DBC,Conn,TempSysRootDir,pic1,rsObj,picSql
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
	if SysRootDir = "" then
		TempSysRootDir = ""
	else
		TempSysRootDir = "/" & SysRootDir
end if
pic1=replace(request("pic"),"'","")
set rsObj=server.createobject(G_FS_RS)
	  picSql="select * from FS_Ads where Location="&pic1&""
	  rsObj.open picSql,Conn,1,1
	  if rsobj.eof then
	     Response.Write"�������ݴ���!!!"
	  else 
%>
      <a href="<%=rsobj("url")%>" target="_blank"><img src="<%=TempSysRootDir%><%=rsobj("LeftPicPath")%>" border="0"></img></a>
<% end if
   rsObj.close
   set rsObj=nothing
 %>
