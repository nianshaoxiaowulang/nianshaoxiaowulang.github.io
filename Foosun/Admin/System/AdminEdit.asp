<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Md5.asp" -->
<%
'==============================================================================
'软件名称：风讯网站信息管理系统
'当前版本：Foosun Content Manager System(FoosunCMS V3.1.0930)
'最新更新：2005.10
'==============================================================================
'Copyright (C) 2002-2004 Foosun.Net  All rights reserved.
'商业注册联系：028-85098980-601,项目开发：028-85098980-606、609,客户支持：608
'产品咨询QQ：394226379,159410,125114015
'技术支持QQ：315485710,66252421 
'项目开发QQ：415637671，655071
'程序开发：四川风讯科技发展有限公司(Foosun Inc.)
'Email:service@Foosun.cn
'MSN：skoolls@hotmail.com
'论坛支持：风讯在线论坛(http://bbs.foosun.net)
'官方网站：www.Foosun.cn  演示站点：test.cooin.com 
'网站通系列(智能快速建站系列)：www.ewebs.cn
'==============================================================================
'免费版本请在程序首页保留版权信息，并做上本站LOGO友情连接
'风讯公司保留此程序的法律追究权利
'如需进行2次开发，必须经过风讯公司书面允许。否则将追究法律责任
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
<title>系统管理员</title>
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
		Response.Write("<script>alert('此管理员可能已经被删除了！');window.close();</script>")
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
			Alert "没有管理员组可供选择，请先添加管理员组" 
		end if
		if AdminName ="" then
			Alert "用户名含有非法字符，请重新输入"  
		end if
		if AdminID = "" then
			RsAdminObj.Open "Select * from FS_Admin where Name='" & AdminName & "'",Conn
		else
			RsAdminObj.Open "Select * from FS_Admin where Name='" & AdminName & "' and ID <>" & AdminID,Conn
		end if
		if Not RsAdminObj.Eof then
			Alert "用户名已经存在"  
		end if
		RsAdminObj.Close
		if AdminID = "" then
			RsAdminObj.Open "Select * from FS_Admin",Conn,3,3
			RsAdminObj.AddNew
		else
			RsAdminObj.Open "Select * from FS_Admin where ID=" & AdminID,Conn,3,3
			if RsAdminObj.Eof then
				Alert= "修改的用户不存在，可能已经被删除"  
			end if
		end if
		if AdminID = "" then
			if Len(AdminPassWord) < 6 then
				Alert "密码至少要六位" 
			end if
			if AdminPassWord <> AffirmPassWord then
				Alert "密码和确认密码不一至"  
			end if
			RsAdminObj("PassWord") = md5(AdminPassWord,16)
			RsAdminObj("RegTime") = Now
		end if
		RsAdminObj("Name") = NoCSSHackAdmin(AdminName,"用户名称")
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
          <td width=35 align="center" alt="保存" onClick="AddSubmit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
		  <td>&nbsp; </td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#DDDDDD">
    <tr> 
      <td width="100" height="26" bgcolor="#E8E8E8"> 
        <div align="right">用户名称 
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
        <div align="right">您的密码</div></td>
      <td bgcolor="#FFFFFF"> 
        <input name="PassWord" type="password" style="width:95%;" id="PassWord2" size="36" value="<% = AdminPassWord %>"> 
        <font color="#FF0000">*</font></td>
    </tr>
    <tr> 
      <td height="26" bgcolor="#E8E8E8"> 
        <div align="right">确认密码</div></td>
      <td bgcolor="#FFFFFF"> 
        <input name="AffirmPassWord" type="password" style="width:95%;" id="AffirmPassWord2" size="36" value="<% = AffirmPassWord %>"> 
        <font color="#FF0000">*</font></td>
    </tr>
		  <% end if %>
    <tr> 
      <td height="26" bgcolor="#E8E8E8"> 
        <div align="right">所属组</div></td>
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
        <div align="right">真实姓名</div></td>
      <td bgcolor="#FFFFFF"> 
        <input style="width:95%;" value="<% =AdminRealName %>" name="RealName" type="text" id="RealName2" size="36" maxlength="50"> 
        <font color="#FF0000">*</font></td>
    </tr>
    <tr> 
      <td height="26" bgcolor="#E8E8E8"> 
        <div align="right">性 别</div></td>
      <td bgcolor="#FFFFFF"> 
        <input <% if AdminSex = 0 then Response.Write("checked") %> name="Sex" type="radio" value="0" checked>
        男 
        <input <% if AdminSex = 1 then Response.Write("checked") %> type="radio" name="Sex" value="1">
        女</td>
    </tr>
    <% if AdminGroupID <> "0" then %>
    <tr> 
      <td height="26" bgcolor="#E8E8E8"> 
        <div align="right">是否锁定</div></td>
      <td bgcolor="#FFFFFF"> 
        <input name="Lock" <% if AdminLock = 1 then Response.write("checked") %> type="checkbox" id="Lock" value="1">
        是否锁定</td>
    </tr>
    <% end if %>
    <tr> 
      <td height="26" bgcolor="#E8E8E8"> 
        <div align="right">邮箱地址</div></td>
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
        <div align="right">主页地址</div></td>
      <td bgcolor="#FFFFFF"> 
        <input style="width:95%;" value="<% =AdminHomePage %>" name="HomePage" type="text" id="HomePage2" size="36" maxlength="150"></td>
    </tr>
    <tr> 
      <td bgcolor="#E8E8E8"> 
        <div align="right">简要说明</div></td>
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
	if (document.AdminForm.Name.value=='') ErrorCode=ErrorCode+'没有填写用户名！\n';
	<% if AdminID = "" then %>
	if (document.AdminForm.PassWord.value=='') ErrorCode=ErrorCode+'没有填写密码！\n';
	if (document.AdminForm.AffirmPassWord.value=='') ErrorCode=ErrorCode+'没有填写确认密码！\n';
	if (document.AdminForm.PassWord.value!=document.AdminForm.AffirmPassWord.value) ErrorCode=ErrorCode+'密码和确认密码不符！\n';
	<% end if %>
	if (document.AdminForm.GroupID.value=='') ErrorCode=ErrorCode+'没有填写管理员组！\n';
	if (document.AdminForm.RealName.value=='') ErrorCode=ErrorCode+'没有填写真实姓名！\n';
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
