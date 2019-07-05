<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
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

%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Request("OrdinaryID") <> "" then
	if Not JudgePopedomTF(Session("Name"),"P070402") then Call ReturnError1()
else
	if Not JudgePopedomTF(Session("Name"),"P070401") then Call ReturnError1()
end if
Dim Result
Dim OrdinaryID,OrdinaryName,OrdinaryUrl,RsTempObj,OperateType,Sql
OrdinaryID = Request("OrdinaryID")
OperateType = Request("OperateType")
Result = Request.Form("Result")
if OrdinaryID <> "" then
	Sql = "Select * from FS_Routine where ID=" & OrdinaryID & " and Type=" & OperateType
	Set RsTempObj = Server.CreateObject(G_FS_RS)
	RsTempObj.Open Sql,Conn,3,3
	if Not RsTempObj.Eof then
		OrdinaryName = RsTempObj("Name")
		OrdinaryUrl = RsTempObj("Url")
		if Result = "Submit" then
			OrdinaryName = Request.Form("Name")
			OrdinaryUrl = Request.Form("Url")
			'Dim OrdNameModObj
			'Set OrdNameModObj = Conn.Execute("Select ID from Routine where Name='"&Request.Form("Name")&"' and Type="&OperateType&"")
			'If Not OrdNameModObj.eof then
			'	Response.Write("<script>alert(""名称与已有记录重复,请重新输入"");</script>")
				'Response.End
			'else
				RsTempObj("Name") = NoCSSHackAdmin(Request.Form("Name"),"名称")
				RsTempObj("Url") = Request.Form("Url")
				RsTempObj.UpDate
				if Err.Number = 0 then
					Response.Redirect("OrdinaryList.asp?Type=" & OperateType)
				else
					%>
					<script language="JavaScript">
						alert('修改失败');
					</script>
					<%
				end if
			'End If
			OrdNameModObj.Close
			Set OrdNameModObj = Nothing
		end if
	else
		%>
		<script language="JavaScript">
			alert('参数传递错误');
		</script>
		<%
	end if
	Set RsTempObj = Nothing
else
	OrdinaryName = ""
    If OperateType="2" or OperateType="5" then
		OrdinaryUrl = "http://"
	Elseif OperateType="3" or OperateType="4" then
		OrdinaryUrl = "mailto:"
	else
		OrdinaryUrl = ""
	End If
	if Result = "Submit" then
		OrdinaryName = NoCSSHackAdmin(Request.Form("Name"),"名称")
		OrdinaryUrl = Request.Form("Url")
		Dim OrdNameObj
		Set OrdNameObj = Conn.Execute("Select ID from FS_Routine where Name='"&Request.Form("Name")&"' and Type="&OperateType&"")
		If Not OrdNameObj.eof then
			Response.Write("<script>alert(""名称与已有记录重复,请重新输入"");</script>")
		else
			Sql = "insert into FS_Routine(Name,Url,Type) values ('" & Request.Form("Name") & "','" & Request.Form("Url") & "'," & OperateType & ")"
			Conn.Execute(Sql)
			if Err.Number = 0 then
				Response.Redirect("OrdinaryList.asp?Type=" & OperateType)
			else
				%>
				<script language="JavaScript">
					alert('添加失败');
				</script>
				<%
			end if
		End If
		OrdNameObj.Close
		Set OrdNameObj = Nothing
	end if
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>常规管理添加和修改</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<link href="Class.css" rel="stylesheet">
<body topmargin="2" leftmargin="2">
<form name="OrdinaryForm" action="" method="post">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
		  <td width=35 align="center" alt="保存" onClick="CheckForm();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
            <td>&nbsp; <input name="OperateType" value="<% = OperateType %>" type="hidden" id="OperateType"> 
              <input name="Result" type="hidden" id="Result" value="Submit"> <input type="hidden" value="<% = OrdinaryID %>" name="OrdinaryID"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="100%" border="0" cellspacing="1" cellpadding="3">
    <tr> 
      <td width="105"><div align="center">名称</div></td>
      <td><input style="width:100%;" value="<% = OrdinaryName %>" type="text" name="Name"></td>
    </tr>
    <tr id="LinkShow"> 
      <td><div align="center">链接</div></td>
      <td><input style="width:100%;" value="<% = OrdinaryUrl %>" type="text" name="Url"></td>
    </tr>
</table>
</form>
</body>
</html>
<script>
var OrdType = '<% = OperateType %>';
function ChooseType()
{
 if (OrdType=='1') document.all.LinkShow.style.display='none';
 else document.OrdinaryForm.Url.disabled=false;
}
function CheckForm()
{
	if (document.OrdinaryForm.Name.value=='')
	{
		alert('请填写名称');
	}
	else
	{
		document.OrdinaryForm.submit();
	}
}
ChooseType();
</script>
<%
Set Conn = Nothing
%>