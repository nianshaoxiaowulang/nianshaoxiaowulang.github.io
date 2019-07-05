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
Dim RsUserConfigObj
Set RsUserConfigObj = Conn.execute("select sitename from FS_Config")
Dim IContent
IContent = Replace(Replace(Request("Content"),"""","%22"),"'","%27")
If  Request.Form("action") = "add" then
	'Response.Write(Request.Form("Content"))
	'Response.end
    Dim UserAddObj,UserAddSql,ChooseMemNameObj,MemNameStr
	If NoCSSHackAdmin(Request.Form("Title"),"标题")="" or isnull(Request.Form("Title")) then
		Response.Write("<script>alert(""请填写标题"");location=""javascript:history.back()"";</script>")
		Response.End
	Else
	End If
	If len(Request.Form("Title"))>100 then
		Response.Write("<script>alert(""标题不可以超过100个字符"");location=""javascript:history.back()"";</script>")
		Response.End
	End If 
	If Request.Form("Content")="" then
		Response.Write("<script>alert(""请填写内容"");location=""javascript:history.back()"";</script>")
		Response.End
	End If
	Set UserAddObj = Server.CreateObject(G_FS_RS)
		UserAddSql = "Select * from FS_MemberNews where 1=0"
		UserAddObj.Open UserAddSql,Conn,3,3
		UserAddObj.AddNew
		UserAddObj("Title") = Replace(Replace(Request.Form("Title"),"""",""),"'","")
		UserAddObj("Content") = Request.Form("Content")
		UserAddObj("Popid") = Cint(Request.Form("Popid"))
		UserAddObj("Author") = Replace(Replace(Request.Form("Author"),"""",""),"'","")
		If Request.Form("isLock") = "0" then
			UserAddObj("isLock") = 0
		Else
			UserAddObj("isLock") = 1
		End If
		UserAddObj("Addtime") = Request.Form("addtime")
		UserAddObj.Update
		UserAddObj.Close
		Set UserAddObj = Nothing
		Response.Redirect("SysUserNews.asp")
		Response.End
End If
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
          <td width=35 align="center" alt="保存"  onClick="SubmitFun();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
          <td>&nbsp;<input name="action" type="hidden" id="action" value="add">
              <input type="hidden" name="Content" value="<% = IContent %>"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="100" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td height="3"></td>
    </tr>
  </table>
  <table width="100%" height="168"  border="0" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
    <tr> 
      <td width="81" height="30" bgcolor="#F2F2F2"> 
        <div align="center">标题　</div></td>
      <td width="909" colspan="3" bgcolor="#F2F2F2"> 
        <input name="Title" type="text"  id="Title" style="width:100%"></td>
    </tr>
    <tr bgcolor="#F2F2F2"> 
      <td height="30"> 
        <div align="center">发布人　</div></td>
      <td colspan="3"> 
        <input name="Author" type="text" id="Author" style="width:100%" value="<% = RsUserConfigObj("SiteName")%>"></td>
    </tr>
    <tr bgcolor="#F2F2F2"> 
      <td height="27"> 
        <div align="center">浏览权限　</div></td>
      <td colspan="3"> 
        <select name="PoPid" id="PoPid">
          <option value="0" selected>所有人</option>
          <option value="1">一般会员</option>
          <option value="2">中级会员</option>
          <option value="3">高级会员</option>
          <option value="4">VIP会员</option>
        </select></td>
    </tr>
    <tr> 
      <td height="19" colspan="4" bgcolor="#EBEBEB"> 
        <iframe id='NewsContent' src='../../Editer/NewsEditer.asp' frameborder=0 scrolling=no width='100%' height='350'></iframe></td>
    </tr>
    <tr bgcolor="#F2F2F2"> 
      <td height="31"> 
        <div align="center">发布时间　</div></td>
      <td colspan="3"> 
        <input name="Addtime" type="text" id="Addtime" value="<% = Now %>">
        ,请正确填写时间格式。</td>
    </tr>
    <tr valign="middle" bgcolor="#F2F2F2"> 
      <td> 
        <div align="center">锁&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;定　</div></td>
      <td> 
        <input type="radio" name="isLock" value="1" <%If Request("Lock") = "1" then Response.Write("checked") end if%>>
        是 
        <input name="isLock" type="radio" value="0" <%If Request("Lock") = "0" or Request("Lock") = "" then Response.Write("checked") end if%>>
        否</td>
    </tr>
  </table>
</form>
</body>
</html>
<%
RsUserConfigObj.Close
Set RsUserConfigObj = Nothing
%>
<script language="JavaScript" type="text/JavaScript">
function SubmitFun()
{
	if (frames["NewsContent"].CurrMode!='EDIT') {alert('其他模式下无法保存，请切换到设计模式');return;}
	frames["NewsContent"].SaveCurrPage();
	var TempContentArray=frames["NewsContent"].NewsContentArray;
	document.UserAddSForm.Content.value='';
	for (var i=0;i<TempContentArray.length;i++)
	{
		if (TempContentArray[i]!='')
		{
			if (document.UserAddSForm.Content.value=='') document.UserAddSForm.Content.value=TempContentArray[i];
			else document.UserAddSForm.Content.value=document.UserAddSForm.Content.value+'[Page]'+TempContentArray[i];
		} 
	}
	document.UserAddSForm.submit();
}</script>
