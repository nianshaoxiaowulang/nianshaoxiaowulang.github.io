<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Md5.asp" -->
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
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P040204") then Call ReturnError1()
if request.Form("action")="add" then
	if len(request.Form("OldPassWord"))<1 then
		Response.Write("<script>alert(""����\n������ԭ���룡"&Copyright&""");location.href=""ChangePwd.asp"";</script>")
		Response.End
	end if
	if request.Form("PassWord")="" then
		Response.Write("<script>alert(""����\n����д������"&Copyright&""");location.href=""ChangePwd.asp"";</script>")
		Response.End
	end if
	if len(request.Form("PassWord"))<6 then
		Response.Write("<script>alert(""����\n���벻������6���ַ�"&Copyright&""");location.href=""ChangePwd.asp"";</script>")
		Response.End
	end if
	if request.Form("PassWord")<>request.Form("AffirmPassWord") then
		Response.Write("<script>alert(""����\n2�����벻��ͬ"&Copyright&""");location.href=""ChangePwd.asp"";</script>")
		Response.End
	end if
	dim Rs,SQL
	set Rs = server.CreateObject (G_FS_RS)
	SQL="select * from FS_admin where id="&cint(request.Form("id"))&" and name='"&request.Form("AdminName")&"'"
	Rs.Open SQL,Conn,3,3
	If Rs("PassWord")=md5((request.Form("OldPassWord")),16) then
		Rs("PassWord")=md5((request.Form("PassWord")),16)
		Rs.update
	Else
		Response.Write("<script>alert(""����\nԭ���벻��ȷ��"&Copyright&""");location.href=""ChangePwd.asp"";</script>")
		Response.End
	End If

	Rs.close
	Set Rs=nothing
	Response.Write("<script>alert(""��ϲ!��\n������ĳɹ�,�����ص�½ҳ��"&Copyright&""");top.location.href=""../Login.asp"";</script>")
	Response.End
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�޸Ĺ���Ա����</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<body topmargin="2" leftmargin="2">
<form action="" method="post" name="PassWordForm">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="����" onClick="document.PassWordForm.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
          <td>&nbsp;
              <input name="AdminName" type="hidden" id="AdminName" value="<%=session("name")%>"> 
              <input name="id" type="hidden" id="id" value="<%=Session("AdminID")%>"> 
              <input name="action" type="hidden" id="action" value="add"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td height="30" colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <div align="center">ԭ 
          �� �� 
          <input name="OldPassWord" type="password" id="PassWord" style="width:60%;">
        </div></td>
    </tr>
    <tr> 
      <td height="30" colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <div align="center">�� 
          �� �� 
          <input name="PassWord" type="password" id="PassWord" style="width:60%;">
        </div></td>
    </tr>
    <tr> 
      <td height="30" colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <div align="center">ȷ������ 
          <input name="AffirmPassWord" type="password" id="AffirmPassWord" style="width:60%;">
        </div></td>
    </tr>
  </table>
</form>
</body>
</html>
