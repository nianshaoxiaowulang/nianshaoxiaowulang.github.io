<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
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
if Not ((JudgePopedomTF(Session("Name"),"P040203")) OR (JudgePopedomTF(Session("Name"),"P040206"))) then Call ReturnError()
Dim AdminID,Result,OperateType,PromptText
PromptText = ""
OperateType = Request("OperateType")
Result = Request("Result")
AdminID = Replace(Replace(Request("AdminID"),"'",""),"""","")
if OperateType = "DelAdmin" then
	if Not JudgePopedomTF(Session("Name"),"P040203") then Call ReturnError()
	PromptText = "确定要删除此管理员吗？"
elseif OperateType = "LockAdmin" then
	if Not JudgePopedomTF(Session("Name"),"P040206") then Call ReturnError()
	PromptText = "确定要锁定此管理员吗？"
elseif OperateType = "UNLockAdmin" then
	if Not JudgePopedomTF(Session("Name"),"P040206") then Call ReturnError()
	PromptText = "确定要解锁此管理员吗？"
else
	PromptText = ""
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>删除或者锁定管理员</title>
</head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet">
<body  oncontextmenu="return false;">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <form name="OperateForm" action="" method="post">
  <tr> 
      <td width="7%" height="20">&nbsp;</td>
      <td width="27%" rowspan="3"><div align="center"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
      <td width="66%">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><% = PromptText %></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td height="20">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="3"><div align="center"><input type="submit" name="Submit" value=" 确 定 ">
          <input name="OperateType" value="<% = OperateType %>" type="hidden" id="OperateType">
          <input name="AdminID" value="<% = AdminID %>" type="hidden" id="AdminID">
          <input name="Result" type="hidden" id="Result" value="Submit">
          <input type="button" name="Submit2" onClick="dialogArguments.location.reload();window.close();" value=" 取 消 ">
      </div></td>
    </tr>
 </form>
</table>
</body>
</html>
<%
if Result = "Submit" then
	Dim ReturnCheckInfo
	AdminID = Replace(AdminID,"***",",")
	if OperateType = "DelAdmin" then
		if Not JudgePopedomTF(Session("Name"),"P040203") then Call ReturnError()
		Conn.Execute("delete from FS_Admin where ID in (" & AdminID & ") and GroupID<>0")
	elseif OperateType = "LockAdmin" then
		if Not JudgePopedomTF(Session("Name"),"P040206") then Call ReturnError()
		Conn.Execute("update FS_Admin set Lock=1 where ID in (" & AdminID & ") and GroupID<>0")
	elseif OperateType = "UNLockAdmin" then
		if Not JudgePopedomTF(Session("Name"),"P040206") then Call ReturnError()
		Conn.Execute("update FS_Admin set Lock=0 where ID in (" & AdminID & ") and GroupID<>0")
	end if
	if Err.Number = 0 then
		%>
		<script language="JavaScript">
		dialogArguments.location.reload();
		window.close();
		</script>
		<%
	else
		%>
		<script language="JavaScript">
		alert('发生错误');
		dialogArguments.location.reload();
		window.close();
		</script>
		<%
	end if
end if
%>