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
Set DBC = Nothing
%>
<!--#include file="Comm/User_Purview.Asp" -->
<%
	Dim RsConfigObj
	Set RsConfigObj = Conn.Execute("Select SiteName,UserConfer,Copyright,isEmail,isChange,IsShop from FS_Config")
	Dim RsUserObj
	Set RsUserObj = Conn.Execute("Select Point,RegTime,UserNo,UserPoint,ShopPoint From FS_Members where MemName = '"& Session("MemName")&"' and Password = '"& Session("MemPassword") &"'")
%>
<HTML><HEAD>
<TITLE><%=RsConfigObj("SiteName")%> >> ��Ա����</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="Css/UserCSS.css" type=text/css  rel=stylesheet>
</HEAD>
<body bgcolor="#F5F5F5" leftmargin="3" topmargin="0">
<fieldset>
<legend></legend>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="86%"><strong><font color="#FF0000">��</font></strong> <font color="#FF0000"><strong><%=Session("MemName")%></strong></font> ��ӭ������<font color="#000000"> 
      <%
Dim NewsSql,GetMessageObj,TotleMessage
NewsSql = "Select * from FS_Message Where MeRead='"& session("memname")&"' and ReadTF=0 and isDelR=0 and IsRecyle=0"
Set GetMessageObj = Server.CreateObject(G_FS_RS)
GetMessageObj.Open NewsSql,Conn,1,1
TotleMessage = GetMessageObj.Recordcount
If TotleMessage=0 then
	Response.Write("<a href=User_Message.asp target=main>����Ϣ(0)</a>")
Else
	Response.Write("<a href=User_Message.asp target=main><font color=red><b>�����¶���Ϣ("&TotleMessage&")</b></font></a>")
End If
%>
<span class="f41">���û����:<font color="#FF0000"><% =  RsUserObj("UserNo") %></font> 
<%
If cint(RsConfigObj("isShop"))=1 then
%>
      �����ý��:
<% =  RsUserObj("UserPoint") %>      �����ѻ���:<% =  RsUserObj("ShopPoint") %>
      <%End If%>
      �� ע��ʱ��: 
      <% =  RsUserObj("RegTime") %>
    </td>
    <td width="14%"><div align="center"><a href="main.asp" target="_top"><font color="#FF0000">�������</font></a> 
        | <a href="Comm/LetOut.asp" target="_top">�˳�</a></div></td>
  </tr>
</table>

</fieldset></body>
</html>
<%
Set Conn=nothing
%>
