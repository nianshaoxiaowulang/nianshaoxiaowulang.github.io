<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
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
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not ((JudgePopedomTF(Session("Name"),"P040202")) OR (JudgePopedomTF(Session("Name"),"P040202"))) then Call ReturnError()
Dim AdminID,Result
Result = Request("Result")
AdminID = Replace(Replace(Request("AdminID"),"'",""),"""","")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ϵͳ����Ա</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<%
Dim SelectShowGroupStr,CheckInfo
Dim AdminName,AdminEmail,AdminOicq,AdminHomePage,AdminSex,AdminLock,AdminRealName,AdminSelfInfo,AdminGroupID
Dim RsAdminObj,ReturnCheckInfo,AffirmPassWord,AdminPassWord
if AdminID <> "" then
	Set RsAdminObj = Conn.Execute("Select * from FS_Admin where ID=" & AdminID)
	if RsAdminObj.Eof then
		Response.Write("<script>alert('�˹���Ա�����Ѿ���ɾ���ˣ�');window.close();</script>")
		Response.End
	else
		if Result = "Submit" then
			AdminName = Request.Form("Name")
			AdminEmail = Request.Form("Email")
			AdminOicq = Request.Form("OICQ")
			AdminHomePage = Request.Form("HomePage")
			AdminSex = Request.Form("Sex")
			AdminLock = Request.Form("Lock")
			AdminRealName = Request.Form("RealName")
			AdminSelfInfo = Request.Form("SelfIntro")
			AdminGroupID = Request.Form("GroupID")
			AffirmPassWord = Request.Form("AffirmPassWord")
			AdminPassWord = Request.Form("AdminPassWord")
			SaveSubmit
		else
			AdminName = RsAdminObj("Name")
			AdminEmail = RsAdminObj("Email")
			AdminOicq = RsAdminObj("OICQ")
			AdminHomePage = RsAdminObj("HomePage")
			AdminSex = RsAdminObj("Sex")
			AdminLock = RsAdminObj("Lock")
			AdminRealName = RsAdminObj("RealName")
			AdminSelfInfo = RsAdminObj("SelfIntro")
			AdminGroupID = RsAdminObj("GroupID")
			AffirmPassWord = RsAdminObj("PassWord")
			AdminPassWord = RsAdminObj("PassWord")
		end if
	end if
else
		AdminName = Request.Form("Name")
		AdminEmail = Request.Form("Email")
		AdminOicq = Request.Form("OICQ")
		AdminHomePage = Request.Form("HomePage")
		AdminSex = Request.Form("Sex")
		AdminLock = Request.Form("Lock")
		AdminRealName = Request.Form("RealName")
		AdminSelfInfo = Request.Form("SelfIntro")
		AdminGroupID = Request.Form("GroupID")
		AffirmPassWord = Request.Form("AffirmPassWord")
		AdminPassWord = Request.Form("AdminPassWord")
		if Result = "Submit" then
			SaveSubmit
		end if
end if
if AdminGroupID <> "0" then
	Set RsAdminObj = Conn.Execute("Select * from FS_AdminGroup")
	SelectShowGroupStr = ""
	do while Not RsAdminObj.Eof
		if Clng(RsAdminObj("ID")) = Clng(AdminGroupID) then
			SelectShowGroupStr = SelectShowGroupStr & "<option selected value=" & RsAdminObj("ID") & ">" & RsAdminObj("GroupName") & "</option>"
		else
			SelectShowGroupStr = SelectShowGroupStr & "<option value=" & RsAdminObj("ID") & ">" & RsAdminObj("GroupName") & "</option>"
		end if
		RsAdminObj.MoveNext
	loop
	RsAdminObj.Close
end if
Set RsAdminObj = Nothing
Sub SaveSubmit()
	if Result = "Submit" then
		AdminID = Replace(Replace(AdminID,"'",""),"""","")
		AdminName = Replace(Replace(AdminName,"'",""),"""","")
		AdminEmail = Replace(Replace(AdminEmail,"'",""),"""","")
		AdminOicq = Replace(Replace(AdminOicq,"'",""),"""","")
		AdminHomePage = Replace(Replace(AdminHomePage,"'",""),"""","")
		AdminSex = Replace(Replace(AdminSex,"'",""),"""","")
		AdminLock = Replace(Replace(AdminLock,"'",""),"""","")
		AdminRealName = Replace(Replace(AdminRealName,"'",""),"""","")
		AdminSelfInfo = Replace(Replace(AdminSelfInfo,"'",""),"""","")
		AdminGroupID = Replace(Replace(AdminGroupID,"'",""),"""","")
		AdminPassWord = Replace(Replace(Request.Form("PassWord"),"'",""),"""","")
		AffirmPassWord = Replace(Replace(Request.Form("AffirmPassWord"),"'",""),"""","")
		'On Error Resume Next
		Set RsAdminObj = Server.CreateObject(G_FS_RS)
		if AdminGroupID ="" then
			Alert "û�й���Ա��ɹ�ѡ��������ӹ���Ա��" 
		end if
		if AdminName ="" then
			Alert "�û������зǷ��ַ�������������"  
		end if
		if AdminID = "" then
			RsAdminObj.Open "Select * from FS_Admin where Name='" & AdminName & "'",Conn
		else
			RsAdminObj.Open "Select * from FS_Admin where Name='" & AdminName & "' and ID <>" & AdminID,Conn
		end if
		if Not RsAdminObj.Eof then
			Alert "�û����Ѿ�����"  
		end if
		RsAdminObj.Close
		if AdminID = "" then
			RsAdminObj.Open "Select * from FS_Admin",Conn,3,3
			RsAdminObj.AddNew
		else
			RsAdminObj.Open "Select * from FS_Admin where ID=" & AdminID,Conn,3,3
			if RsAdminObj.Eof then
				Alert= "�޸ĵ��û������ڣ������Ѿ���ɾ��"  
			end if
		end if
		if AdminID = "" then
			if Len(AdminPassWord) < 6 then
				Alert "��������Ҫ��λ" 
			end if
			if AdminPassWord <> AffirmPassWord then
				Alert "�����ȷ�����벻һ��"  
			end if
			RsAdminObj("PassWord") = md5(AdminPassWord,16)
			RsAdminObj("RegTime") = Now
		end if
		RsAdminObj("Name") = NoCSSHackAdmin(AdminName,"�û�����")
		RsAdminObj("Email") = AdminEmail
		RsAdminObj("Oicq") = AdminOicq
		RsAdminObj("HomePage") = AdminHomePage
		RsAdminObj("Sex") = AdminSex
		if AdminLock = "" then
			RsAdminObj("Lock") = 0
		else
			RsAdminObj("Lock") = 1
		end if
		RsAdminObj("RealName") = AdminRealName
		RsAdminObj("SelfIntro") = AdminSelfInfo
		RsAdminObj("GroupID") = AdminGroupID
		RsAdminObj.Update
		if ReturnCheckInfo = "" then
			Response.Redirect("SysAdminList.asp")
		else
			%>
			<script>alert('<% = ReturnCheckInfo %>');history.back();</script>
			<%
		end if
	end if
End Sub
Sub Alert(Str)
	Set RsAdminObj = Nothing
	%>
	<script>alert('<% = Str %>');history.back();</script>
	<%
	Response.End
End Sub
%>
<body scrolling=no leftmargin="2" topmargin="2">
<form action="" id="AdminForm" method="post" name="AdminForm">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="����" onClick="AddSubmit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td>&nbsp; </td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#DDDDDD">
    <tr> 
      <td width="100" height="26" bgcolor="#E8E8E8"> 
        <div align="right">�û����� 
          <input name="Result" type="hidden" id="Result2" value="Submit">
        </div></td>
      <td bgcolor="#FFFFFF"> 
        <input value="<% =AdminName %>" <% if AdminName <> "" then Response.Write("readonly") %> name="Name" style="width:95%;" type="text" id="Name2" size="36" maxlength="40"> 
        <font color="#FF0000">*</font> <input value="<% =AdminID %>" name="AdminID" type="hidden" id="AdminID2"> 
      </td>
    </tr>
    <% if AdminGroupID <> "0" then %>
		<% If AdminID = "" then %>
    <tr> 
      <td height="26" bgcolor="#E8E8E8"> 
        <div align="right">��������</div></td>
      <td bgcolor="#FFFFFF"> 
        <input name="PassWord" type="password" style="width:95%;" id="PassWord2" size="36" value="<% = AdminPassWord %>"> 
        <font color="#FF0000">*</font></td>
    </tr>
    <tr> 
      <td height="26" bgcolor="#E8E8E8"> 
        <div align="right">ȷ������</div></td>
      <td bgcolor="#FFFFFF"> 
        <input name="AffirmPassWord" type="password" style="width:95%;" id="AffirmPassWord2" size="36" value="<% = AffirmPassWord %>"> 
        <font color="#FF0000">*</font></td>
    </tr>
		  <% end if %>
    <tr> 
      <td height="26" bgcolor="#E8E8E8"> 
        <div align="right">������</div></td>
      <td bgcolor="#FFFFFF"> 
        <select style="width:95%;" name="GroupID" id="select">
          <% =SelectShowGroupStr %>
        </select> <font color="#FF0000">*</font></td>
    </tr>
    <% else %>
    <input value="0" type="hidden" name="GroupID">
    <% end if %>
    <tr> 
      <td height="26" bgcolor="#E8E8E8"> 
        <div align="right">��ʵ����</div></td>
      <td bgcolor="#FFFFFF"> 
        <input style="width:95%;" value="<% =AdminRealName %>" name="RealName" type="text" id="RealName2" size="36" maxlength="50"> 
        <font color="#FF0000">*</font></td>
    </tr>
    <tr> 
      <td height="26" bgcolor="#E8E8E8"> 
        <div align="right">�� ��</div></td>
      <td bgcolor="#FFFFFF"> 
        <input <% if AdminSex = 0 then Response.Write("checked") %> name="Sex" type="radio" value="0" checked>
        �� 
        <input <% if AdminSex = 1 then Response.Write("checked") %> type="radio" name="Sex" value="1">
        Ů</td>
    </tr>
    <% if AdminGroupID <> "0" then %>
    <tr> 
      <td height="26" bgcolor="#E8E8E8"> 
        <div align="right">�Ƿ�����</div></td>
      <td bgcolor="#FFFFFF"> 
        <input name="Lock" <% if AdminLock = 1 then Response.write("checked") %> type="checkbox" id="Lock" value="1">
        �Ƿ�����</td>
    </tr>
    <% end if %>
    <tr> 
      <td height="26" bgcolor="#E8E8E8"> 
        <div align="right">�����ַ</div></td>
      <td bgcolor="#FFFFFF"> 
        <input style="width:95%;" value="<% =AdminEmail %>" name="Email" type="text" id="Email2" size="36" maxlength="50"></td>
    </tr>
    <tr> 
      <td height="26" bgcolor="#E8E8E8"> 
        <div align="right">Oicq</div></td>
      <td bgcolor="#FFFFFF"> 
        <input style="width:95%;" value="<% =AdminOicq %>" name="Oicq" type="text" id="Oicq2" size="36" maxlength="15"></td>
    </tr>
    <tr> 
      <td height="26" bgcolor="#E8E8E8"> 
        <div align="right">��ҳ��ַ</div></td>
      <td bgcolor="#FFFFFF"> 
        <input style="width:95%;" value="<% =AdminHomePage %>" name="HomePage" type="text" id="HomePage2" size="36" maxlength="150"></td>
    </tr>
    <tr> 
      <td bgcolor="#E8E8E8"> 
        <div align="right">��Ҫ˵��</div></td>
      <td bgcolor="#FFFFFF"> 
        <textarea style="width:95%;" name="SelfIntro" cols="34" rows="6" id="textarea"><% =AdminSelfInfo %></textarea>
      </td>
    </tr>
</table>
</form>
</body>
</html>
<%
Set Conn = Nothing
%>
<script>
function AddSubmit()
{
	if (CheckAdminForm())
	{
		document.AdminForm.submit();
	}
}
function CheckAdminForm()
{
	var ErrorCode='';
	if (document.AdminForm.Name.value=='') ErrorCode=ErrorCode+'û����д�û�����\n';
	<% if AdminID = "" then %>
	if (document.AdminForm.PassWord.value=='') ErrorCode=ErrorCode+'û����д���룡\n';
	if (document.AdminForm.AffirmPassWord.value=='') ErrorCode=ErrorCode+'û����дȷ�����룡\n';
	if (document.AdminForm.PassWord.value!=document.AdminForm.AffirmPassWord.value) ErrorCode=ErrorCode+'�����ȷ�����벻����\n';
	<% end if %>
	if (document.AdminForm.GroupID.value=='') ErrorCode=ErrorCode+'û����д����Ա�飡\n';
	if (document.AdminForm.RealName.value=='') ErrorCode=ErrorCode+'û����д��ʵ������\n';
	if (ErrorCode!='') 
	{
		alert(ErrorCode);
		return false
	}
	else return true;
}
function SetEmptyForm()
{
	var i;
	for(i=0;i<document.AdminForm.elements.length;i++)
	{
		if (document.AdminForm.elements.item(i).tagName.toLowerCase()=='input')
		{
			if (document.AdminForm.elements.item(i).type=='text') document.AdminForm.elements.item(i).value='';
			if (document.AdminForm.elements.item(i).type=='checkbox') document.AdminForm.elements.item(i).checked=false;
		}
		if (document.AdminForm.elements.item(i).tagName.toLowerCase()=='textarea') document.AdminForm.elements.item(i).innerText='';
	}
}
</script>
