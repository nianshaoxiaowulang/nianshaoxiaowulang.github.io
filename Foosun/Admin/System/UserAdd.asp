<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../../../Inc/Md5.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
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
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P040401") then Call ReturnError1()
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet" type="text/css">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��ӻ�Ա</title>
</head>
<body leftmargin="2" topmargin="2">
<form action="" method="post" name="UserAddSForm">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="����" onClick="document.UserAddSForm.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
          <td>&nbsp;<input name="action" type="hidden" id="action" value="add"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="100%"  border="0" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
    <tr bgcolor="#FFFFFF"> 
      <td width="100" bgcolor="#EBEBEB"> 
        <div align="right">�� Ա ��</div></td>
      <td colspan="3"> 
        <input name="MemName" type="text"  id="MemName" style="width:100%" value="<%=Request("MemName")%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td bgcolor="#EBEBEB"> 
        <div align="right">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</div></td>
      <td colspan="3"> 
        <input name="Password" type="password" id="Password" style="width:1090%"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td bgcolor="#EBEBEB"> 
        <div align="right">ȷ������</div></td>
      <td colspan="3"> 
        <input name="PasswordTF" type="password" id="PasswordTF" style="width:100%"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td bgcolor="#EBEBEB"> 
        <div align="right">�� Ա ��</div></td>
      <td colspan="3"> 
        <select name="GroupID" id="GroupID" style="width:100%">
          <option value="0" <%If Request("GroupID") = "" or  Request("GroupID") = "0" then Response.Write("selected") end if%>> 
          </option>
          <%
		Dim SelGroupObj
		Set SelGroupObj = Conn.Execute("Select GroupID,Name from FS_MemGroup order by PopLevel desc")
		do while not SelGroupObj.eof 
	%>
          <option value="<%=SelGroupObj("GroupID")%>" <%If Cstr(Request("GroupID"))=Cstr(SelGroupObj("GroupID")) then Response.Write("selected") end if%>><%=SelGroupObj("Name")%></option>
          <%
		SelGroupObj.MoveNext
		Loop
		SelGroupObj.Close
		Set SelGroupObj = Nothing
	%>
        </select></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td bgcolor="#EBEBEB"> 
        <div align="right">��ʵ����</div></td>
      <td colspan="3"> 
        <input name="Name" type="text" id="Name" size="20" style="width:100%" value="<%=Request("Name")%>"></td>
    </tr>
    <tr valign="middle" bgcolor="#FFFFFF"> 
      <td bgcolor="#EBEBEB"> 
        <div align="right">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</div></td>
      <td> 
        <input type="radio" name="Lock" value="1" <%If Request("Lock") = "1" then Response.Write("checked") end if%>>
        �� 
        <input name="Lock" type="radio" value="0" <%If Request("Lock") = "0" or Request("Lock") = "" then Response.Write("checked") end if%>>
        ��</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td bgcolor="#EBEBEB"> 
        <div align="right">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</div></td>
      <td> 
        <input name="Sex" type="radio" value="0" <%If Request("Sex") = "0" or Request("Sex") = "" then Response.Write("checked") end if%>>
        �� 
        <input type="radio" name="Sex" value="1" <%If Request("Sex") = "1" then Response.Write("checked") end if%>>
        Ů</td>
    </tr>
</table>
</form>
</body>
</html>
<%
If  Request.Form("action") = "add" then
    Dim UserAddObj,UserAddSql,ChooseMemNameObj,MemNameStr
	If NoCSSHackAdmin(Request.Form("MemName"),"��Ա��")="" or isnull(Request.Form("MemName")) then
		Response.Write("<script>alert(""����д��Ա��¼��"");</script>")
		Response.End
	Else
	End If
///////////////// lzp
	If len(Request.Form("MemName"))>10 then
		Response.Write("<script>alert(""��Ա��¼�������Գ���10���ַ�"");</script>")
		Response.End
	Else
	end if
////////////////
		MemNameStr = Replace(Replace(Request.Form("MemName"),"""",""),"'","")
	
	Set ChooseMemNameObj = Conn.Execute("Select ID from FS_Members where MemName='"&MemNameStr&"'")
	If Not ChooseMemNameObj.eof then
		Response.Write("<script>alert(""�˻�Ա��¼���Ѿ�����,���޸�"");</script>")
		Response.End
	End If
	ChooseMemNameObj.Close
	Set ChooseMemNameObj = Nothing
	If Request.Form("Password")="" or isnull("Password") then
		Response.Write("<script>alert(""�������Ա��¼����"");</script>")
		Response.End
	End If
	If Len(Request.Form("Password")) < 6 then
		Response.Write("<script>alert(""��Ա��¼���벻��������λ"");</script>")
		Response.End
	End If
	If Cstr(Request.Form("Password"))<>Cstr(Request.Form("PasswordTF")) then
		Response.Write("<script>alert(""������ȷ�����벻ͬ"");</script>")
		Response.End
	End If
	If Request.Form("Name")="" or isnull(Request.Form("Name")) then
		Response.Write("<script>alert(""����д��Ա��ʵ����"");</script>")
		Response.End
	End If
	Set UserAddObj = Server.CreateObject(G_FS_RS)
		UserAddSql = "Select * from FS_Members where 1=0"
		UserAddObj.Open UserAddSql,Conn,3,3
		UserAddObj.AddNew
		UserAddObj("MemName") = Replace(Replace(Request.Form("MemName"),"""",""),"'","")
		UserAddObj("Password") = md5(Request.Form("Password"),16)
		UserAddObj("GroupID") = Request.Form("GroupID")
		UserAddObj("Name") = Replace(Replace(Request.Form("Name"),"""",""),"'","")
		If Request.Form("Lock") = "0" then
			UserAddObj("Lock") = "0"
		Else
			UserAddObj("Lock") = "1"
		End If
		If Request.Form("Sex") = "0" then
			UserAddObj("Sex") = "0"
		Else
			UserAddObj("Sex") = "1"
		End If
		UserAddObj("RegTime") = Now()
		UserAddObj("Email") = "foosun@foosun.cn"
		UserAddObj("LastLoginIP") = Request.ServerVariables("REMOTE_ADDR")
		UserAddObj("LastLoginTime") = Now()
		UserAddObj.Update
		UserAddObj.Close
		Set UserAddObj = Nothing
		Response.Redirect("SysUserList.asp")
		Response.End
End If
%>