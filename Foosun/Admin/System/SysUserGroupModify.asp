<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
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
if Not JudgePopedomTF(Session("Name"),"P040302") then Call ReturnError()
    Dim UserGroupID,UserGroupObj
	UserGroupID = Request("ID")
	If Request("ID")="" or isnull(Request("ID")) then
		Response.Write("<script>alert(""参数传递错误"");</script>")
		Response.Redirect("SysUserGroup.asp")
		Response.End
	else
		Set UserGroupObj = Conn.Execute("Select * from FS_MemGroup where ID="&UserGroupID&"")
		if UserGroupObj.eof then
		   Response.Write("<script>alert(""参数传递错误"");</script>")
			Response.Redirect("SysUserGroup.asp")
		   Response.End
		end if
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>会员组修改</title>
</head>
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<body leftmargin="2" topmargin="2">
<form action="" name="UserGroupForm" method="post">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="保存" onClick="document.UserGroupForm.submit();;" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
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
        <div align="center">组&nbsp;&nbsp;&nbsp;&nbsp;名</div></td>
      <td> 
        <input name="Name" type="text" id="Name" style="width:100%" title="会员组名称,长度不能超过25个中文字符" value="<%=UserGroupObj("Name")%>" maxlength="25"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26"> 
        <div align="center">权限级别</div></td>
      <td> 
        <input name="PopLevel" type="text" id="PopLevel" style="width:100%" title="输入会员权限级别,数值越小,权限越大,范围:10-32767,用于设置新闻浏览权限" onBlur="CheckNumber(this,'权限级别');" value="<%=UserGroupObj("PopLevel")%>" maxlength="9"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26"> 
        <div align="center">初始积分</div></td>
      <td> 
        <input name="Point" type="text" id="Point" style="width:100%" title="会员的初始积分数,分数越大,权限越高,多用于新闻浏览以外的其它会员动作" onBlur="CheckNumber(this,'初始积分');" value="<%=UserGroupObj("Point")%>" maxlength="9"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td valign="middle"> 
        <div align="center">说&nbsp;&nbsp;&nbsp;&nbsp;明</div></td>
      <td> 
        <textarea name="Comment" rows="6" id="Comment" style="width:100%" title="对会员组的说明文字,方便后台管理"><%=UserGroupObj("Comment")%></textarea></td>
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
			UserGroupObj("Name") = NoCSSHackAdmin(Replace(Replace(Request.Form("Name"),"""",""),"'",""),"组名")
		Else
			Response.Write("<script>alert(""请输入会员组名"");</script>")
			Response.End
		End If
		' 待处理
		If  Request.Form("PopLevel")<>"" then
		    If Isnumeric(Request.Form("PopLevel")) and Request.Form("PopLevel")>10 and Request.Form("PopLevel")<32767 then
				UserGroupObj("PopLevel") = Cint(Request.Form("PopLevel"))
			Else
				Response.Write("<script>alert(""会员组级别必须为数字类型,且不能小于10大于32767"");</script>")
				Response.End
			End If
		Else
			Response.Write("<script>alert(""会员组级别必须为数字类型,且不能小于10大于32767"");</script>")
			Response.End
		End If
		If Request.Form("Comment")<>"" then
			UserGroupObj("Comment") = Request.Form("Comment")
		End If
		UserGroupObj("SysTypeTF") = "0"
		If Isnumeric(Request.Form("Point")) then
			UserGroupObj("Point") = Request.Form("Point")
		Else
			Response.Write("<script>alert(""会员组级别必须为数字类型"");</script>")
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