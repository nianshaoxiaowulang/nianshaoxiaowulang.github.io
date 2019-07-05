<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../../../Inc/Md5.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
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
if Not JudgePopedomTF(Session("Name"),"P040401") then Call ReturnError1()
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet" type="text/css">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>添加会员</title>
</head>
<body leftmargin="2" topmargin="2">
<form action="" method="post" name="UserAddSForm">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="保存" onClick="document.UserAddSForm.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
          <td>&nbsp;<input name="action" type="hidden" id="action" value="add"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="100%"  border="0" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
    <tr bgcolor="#FFFFFF"> 
      <td width="100" bgcolor="#EBEBEB"> 
        <div align="right">会 员 名</div></td>
      <td colspan="3"> 
        <input name="MemName" type="text"  id="MemName" style="width:100%" value="<%=Request("MemName")%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td bgcolor="#EBEBEB"> 
        <div align="right">密&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;码</div></td>
      <td colspan="3"> 
        <input name="Password" type="password" id="Password" style="width:1090%"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td bgcolor="#EBEBEB"> 
        <div align="right">确认密码</div></td>
      <td colspan="3"> 
        <input name="PasswordTF" type="password" id="PasswordTF" style="width:100%"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td bgcolor="#EBEBEB"> 
        <div align="right">会 员 组</div></td>
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
        <div align="right">真实姓名</div></td>
      <td colspan="3"> 
        <input name="Name" type="text" id="Name" size="20" style="width:100%" value="<%=Request("Name")%>"></td>
    </tr>
    <tr valign="middle" bgcolor="#FFFFFF"> 
      <td bgcolor="#EBEBEB"> 
        <div align="right">锁&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;定</div></td>
      <td> 
        <input type="radio" name="Lock" value="1" <%If Request("Lock") = "1" then Response.Write("checked") end if%>>
        是 
        <input name="Lock" type="radio" value="0" <%If Request("Lock") = "0" or Request("Lock") = "" then Response.Write("checked") end if%>>
        否</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td bgcolor="#EBEBEB"> 
        <div align="right">性&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;别</div></td>
      <td> 
        <input name="Sex" type="radio" value="0" <%If Request("Sex") = "0" or Request("Sex") = "" then Response.Write("checked") end if%>>
        男 
        <input type="radio" name="Sex" value="1" <%If Request("Sex") = "1" then Response.Write("checked") end if%>>
        女</td>
    </tr>
</table>
</form>
</body>
</html>
<%
If  Request.Form("action") = "add" then
    Dim UserAddObj,UserAddSql,ChooseMemNameObj,MemNameStr
	If NoCSSHackAdmin(Request.Form("MemName"),"会员名")="" or isnull(Request.Form("MemName")) then
		Response.Write("<script>alert(""请填写会员登录名"");</script>")
		Response.End
	Else
	End If
///////////////// lzp
	If len(Request.Form("MemName"))>10 then
		Response.Write("<script>alert(""会员登录名不可以超过10个字符"");</script>")
		Response.End
	Else
	end if
////////////////
		MemNameStr = Replace(Replace(Request.Form("MemName"),"""",""),"'","")
	
	Set ChooseMemNameObj = Conn.Execute("Select ID from FS_Members where MemName='"&MemNameStr&"'")
	If Not ChooseMemNameObj.eof then
		Response.Write("<script>alert(""此会员登录名已经存在,请修改"");</script>")
		Response.End
	End If
	ChooseMemNameObj.Close
	Set ChooseMemNameObj = Nothing
	If Request.Form("Password")="" or isnull("Password") then
		Response.Write("<script>alert(""请输入会员登录密码"");</script>")
		Response.End
	End If
	If Len(Request.Form("Password")) < 6 then
		Response.Write("<script>alert(""会员登录密码不能少于六位"");</script>")
		Response.End
	End If
	If Cstr(Request.Form("Password"))<>Cstr(Request.Form("PasswordTF")) then
		Response.Write("<script>alert(""密码与确认密码不同"");</script>")
		Response.End
	End If
	If Request.Form("Name")="" or isnull(Request.Form("Name")) then
		Response.Write("<script>alert(""请填写会员真实姓名"");</script>")
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