<% Option Explicit %>
<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Const.asp" -->
<!--#include file="../../Inc/Md5.asp" -->
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
Dim  DBC,Conn
Set  DBC = New DataBaseClass
Set  Conn = DBC.OpenConnection()
Set  DBC = Nothing
Dim Pid
Pid=Replace(Replace(request("Pid"),"'",""),Chr(39),"")
Conn.execute("update FS_Shop_Products set ClickNum=ClickNum+1 where id="&Pid)
Dim Rs
Set Rs = server.createobject(G_FS_RS)
Rs.source = "select ClickNum from FS_Shop_Products where id="&pid
Rs.open rs.source,conn,1,1
If Not Rs.Eof then
%>
   javastr="<%=rs("ClickNum")%>"
   document.write(javastr)
<%
else
%>
   javastr="0"
   document.write(javastr)
<%
End if
Rs.close
set Rs=nothing
Set Conn = Nothing
%>
