<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
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
if Not JudgePopedomTF(Session("Name"),"P040302") then Call ReturnError()
    Dim UserGroupID,UserGroupObj
	UserGroupID = Request("ID")
	If Request("ID")="" or isnull(Request("ID")) then
		Response.Write("<script>alert(""�������ݴ���"");</script>")
		Response.Redirect("SysUserGroup.asp")
		Response.End
	else
		Set UserGroupObj = Conn.Execute("Select * from FS_MemGroup where ID="&UserGroupID&"")
		if UserGroupObj.eof then
		   Response.Write("<script>alert(""�������ݴ���"");</script>")
			Response.Redirect("SysUserGroup.asp")
		   Response.End
		end if
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��Ա���޸�</title>
</head>
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<body leftmargin="2" topmargin="2">
<form action="" name="UserGroupForm" method="post">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="����" onClick="document.UserGroupForm.submit();;" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
          <td>&nbsp;
              <input name="action" type="hidden" id="action" value="mod"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#EBEBEB">
    <tr bgcolor="#FFFFFF"> 
      <td width="100" height="26"> 
        <div align="center">��&nbsp;&nbsp;&nbsp;&nbsp;��</div></td>
      <td> 
        <input name="Name" type="text" id="Name" style="width:100%" title="��Ա������,���Ȳ��ܳ���25�������ַ�" value="<%=UserGroupObj("Name")%>" maxlength="25"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26"> 
        <div align="center">Ȩ�޼���</div></td>
      <td> 
        <input name="PopLevel" type="text" id="PopLevel" style="width:100%" title="�����ԱȨ�޼���,��ֵԽС,Ȩ��Խ��,��Χ:10-32767,���������������Ȩ��" onBlur="CheckNumber(this,'Ȩ�޼���');" value="<%=UserGroupObj("PopLevel")%>" maxlength="9"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26"> 
        <div align="center">��ʼ����</div></td>
      <td> 
        <input name="Point" type="text" id="Point" style="width:100%" title="��Ա�ĳ�ʼ������,����Խ��,Ȩ��Խ��,������������������������Ա����" onBlur="CheckNumber(this,'��ʼ����');" value="<%=UserGroupObj("Point")%>" maxlength="9"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td valign="middle"> 
        <div align="center">˵&nbsp;&nbsp;&nbsp;&nbsp;��</div></td>
      <td> 
        <textarea name="Comment" rows="6" id="Comment" style="width:100%" title="�Ի�Ա���˵������,�����̨����"><%=UserGroupObj("Comment")%></textarea></td>
    </tr>
</table>
</form>
</body>
</html>
<%
UserGroupObj.Close
Set UserGroupObj = Nothing
 If Request.Form("action")="mod" then
    Dim UserGroupSql
	Set UserGroupObj=server.createobject(G_FS_RS)
		UserGroupSql="select * from FS_MemGroup where ID=" & UserGroupID & ""
		UserGroupObj.open UserGroupSql,Conn,3,3
		If Request.Form("Name") <> "" then
			UserGroupObj("Name") = NoCSSHackAdmin(Replace(Replace(Request.Form("Name"),"""",""),"'",""),"����")
		Else
			Response.Write("<script>alert(""�������Ա����"");</script>")
			Response.End
		End If
		' ������
		If  Request.Form("PopLevel")<>"" then
		    If Isnumeric(Request.Form("PopLevel")) and Request.Form("PopLevel")>10 and Request.Form("PopLevel")<32767 then
				UserGroupObj("PopLevel") = Cint(Request.Form("PopLevel"))
			Else
				Response.Write("<script>alert(""��Ա�鼶�����Ϊ��������,�Ҳ���С��10����32767"");</script>")
				Response.End
			End If
		Else
			Response.Write("<script>alert(""��Ա�鼶�����Ϊ��������,�Ҳ���С��10����32767"");</script>")
			Response.End
		End If
		If Request.Form("Comment")<>"" then
			UserGroupObj("Comment") = Request.Form("Comment")
		End If
		UserGroupObj("SysTypeTF") = "0"
		If Isnumeric(Request.Form("Point")) then
			UserGroupObj("Point") = Request.Form("Point")
		Else
			Response.Write("<script>alert(""��Ա�鼶�����Ϊ��������"");</script>")
			Response.End
		End If
		UserGroupObj.Update
		UserGroupObj.Close
		Set UserGroupObj = Nothing
		Response.Redirect("SysUserGroup.asp")
 End If
end if
Conn.Close
Set Conn = Nothing
%>