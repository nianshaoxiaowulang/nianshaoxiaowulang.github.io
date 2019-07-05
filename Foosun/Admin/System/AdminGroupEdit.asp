<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp"-->
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
			AdminGroupObj("GroupName") = NoCSSHackAdmin(Request.Form("GroupName"),"组名称")
			AdminGroupObj("Comment") = Request.Form("Comment")
			AdminGroupObj.UpDate
			if Err.Number = 0 then
				Response.Redirect("SysAdminGroup.asp")
			else
				%>
				<script language="JavaScript">
					alert('修改失败');
				</script>
				<%
			end if
		end if
		GroupName = AdminGroupObj("GroupName")
		Comment = AdminGroupObj("Comment")
	else
		%>
		<script language="JavaScript">
			alert('参数传递错误');
		</script>
		<%
	end if
	Set AdminGroupObj = Nothing
else
	GroupName = ""
	Comment = ""
	if Result = "Submit" then
		if NoCSSHackAdmin(Request.Form("GroupName"),"组名称") <> "" then
			Sql = "Insert into FS_AdminGroup(GroupName,Comment) values ('" & Request.Form("GroupName") & "','" & Request.Form("Comment") & "')"
			Conn.Execute(Sql)
			if Err.Number = 0 then
				Response.Redirect("SysAdminGroup.asp")
			else
				%>
				<script language="JavaScript">
					alert('添加失败');
				</script>
				<%
			end if
		else
			%>
			<script language="JavaScript">
				alert('请填写组名');
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
<title>添加和修改系统管理员组</title>
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
          <td width=35 align="center" alt="保存" onClick="document.AdminGroupForm.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
          <td>&nbsp;<input name="Result" type="hidden" id="Result" value="Submit"> <input type="hidden" value="<% = ID %>" name="OrdinaryID"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr> 
      <td width="100">
<div align="center">组 名 称</div></td>
      <td> <input value="<% =GroupName %>" name="GroupName" style="width:100%;" type="text"  size="36" maxlength="40"></td>
    </tr>
    <tr> 
      <td> <div align="center">简要说明</div></td>
      <td> <textarea style="width:100%;" name="Comment" rows="6" id="textarea"><% = Comment %></textarea></td>
    </tr>
</table>		
</form>
</body>
</html>
<%
Set Conn = Nothing
%>