<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp"-->
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
if Not ((JudgePopedomTF(Session("Name"),"P040101")) OR (JudgePopedomTF(Session("Name"),"P040102"))) then Call ReturnError1()
Dim Result
Dim ID,GroupName,Comment,AdminGroupObj,Sql
ID = Request("ID")
Result = Request.Form("Result")
if ID <> "" then
	Sql = "Select * from FS_AdminGroup where ID=" & ID
	Set AdminGroupObj = Server.CreateObject(G_FS_RS)
	AdminGroupObj.Open Sql,Conn,3,3
	if Not AdminGroupObj.Eof then
		if Result = "Submit" then
			AdminGroupObj("GroupName") = NoCSSHackAdmin(Request.Form("GroupName"),"������")
			AdminGroupObj("Comment") = Request.Form("Comment")
			AdminGroupObj.UpDate
			if Err.Number = 0 then
				Response.Redirect("SysAdminGroup.asp")
			else
				%>
				<script language="JavaScript">
					alert('�޸�ʧ��');
				</script>
				<%
			end if
		end if
		GroupName = AdminGroupObj("GroupName")
		Comment = AdminGroupObj("Comment")
	else
		%>
		<script language="JavaScript">
			alert('�������ݴ���');
		</script>
		<%
	end if
	Set AdminGroupObj = Nothing
else
	GroupName = ""
	Comment = ""
	if Result = "Submit" then
		if NoCSSHackAdmin(Request.Form("GroupName"),"������") <> "" then
			Sql = "Insert into FS_AdminGroup(GroupName,Comment) values ('" & Request.Form("GroupName") & "','" & Request.Form("Comment") & "')"
			Conn.Execute(Sql)
			if Err.Number = 0 then
				Response.Redirect("SysAdminGroup.asp")
			else
				%>
				<script language="JavaScript">
					alert('���ʧ��');
				</script>
				<%
			end if
		else
			%>
			<script language="JavaScript">
				alert('����д����');
			</script>
			<%
			GroupName = Request.Form("GroupName")
			Comment = Request.Form("Comment")
		end if
	end if
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��Ӻ��޸�ϵͳ����Ա��</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<body topmargin="2" leftmargin="2">
<form action="" method="post" name="AdminGroupForm">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="����" onClick="document.AdminGroupForm.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
          <td>&nbsp;<input name="Result" type="hidden" id="Result" value="Submit"> <input type="hidden" value="<% = ID %>" name="OrdinaryID"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr> 
      <td width="100">
<div align="center">�� �� ��</div></td>
      <td> <input value="<% =GroupName %>" name="GroupName" style="width:100%;" type="text"  size="36" maxlength="40"></td>
    </tr>
    <tr> 
      <td> <div align="center">��Ҫ˵��</div></td>
      <td> <textarea style="width:100%;" name="Comment" rows="6" id="textarea"><% = Comment %></textarea></td>
    </tr>
</table>		
</form>
</body>
</html>
<%
Set Conn = Nothing
%>