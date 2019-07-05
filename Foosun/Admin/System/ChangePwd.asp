<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
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
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P040204") then Call ReturnError1()
if request.Form("action")="add" then
	if len(request.Form("OldPassWord"))<1 then
		Response.Write("<script>alert(""错误：\n请输入原密码！"&Copyright&""");location.href=""ChangePwd.asp"";</script>")
		Response.End
	end if
	if request.Form("PassWord")="" then
		Response.Write("<script>alert(""错误：\n请填写新密码"&Copyright&""");location.href=""ChangePwd.asp"";</script>")
		Response.End
	end if
	if len(request.Form("PassWord"))<6 then
		Response.Write("<script>alert(""错误：\n密码不能少于6个字符"&Copyright&""");location.href=""ChangePwd.asp"";</script>")
		Response.End
	end if
	if request.Form("PassWord")<>request.Form("AffirmPassWord") then
		Response.Write("<script>alert(""错误：\n2次密码不相同"&Copyright&""");location.href=""ChangePwd.asp"";</script>")
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
		Response.Write("<script>alert(""错误：\n原密码不正确！"&Copyright&""");location.href=""ChangePwd.asp"";</script>")
		Response.End
	End If

	Rs.close
	Set Rs=nothing
	Response.Write("<script>alert(""恭喜!：\n密码更改成功,将返回登陆页面"&Copyright&""");top.location.href=""../Login.asp"";</script>")
	Response.End
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>修改管理员密码</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<body topmargin="2" leftmargin="2">
<form action="" method="post" name="PassWordForm">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="保存" onClick="document.PassWordForm.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
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
      <td height="30" colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <div align="center">原 
          密 码 
          <input name="OldPassWord" type="password" id="PassWord" style="width:60%;">
        </div></td>
    </tr>
    <tr> 
      <td height="30" colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <div align="center">新 
          密 码 
          <input name="PassWord" type="password" id="PassWord" style="width:60%;">
        </div></td>
    </tr>
    <tr> 
      <td height="30" colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <div align="center">确认密码 
          <input name="AffirmPassWord" type="password" id="AffirmPassWord" style="width:60%;">
        </div></td>
    </tr>
  </table>
</form>
</body>
</html>
