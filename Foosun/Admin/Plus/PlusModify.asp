<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P080602") then Call ReturnError1()
    Dim PlusID,PlusObj
	PlusID = Request("PlusID")
	Set PlusObj = Conn.Execute("Select * from FS_Plus where ID="&PlusID&"")
	If PlusObj.eof then
	   Response.Write("<script>alert(""参数传递错误"");dialogArguments.location.reload();window.close();</script>")
	   Response.End
	End If
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet" type="text/css">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>修改插件</title>
</head>
<body leftmargin="2" topmargin="2">
<form action="" method="post" name="PlusForm">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="保存" onClick="document.PlusForm.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
          <td>&nbsp;<input name="action" type="hidden" id="action" value="mod"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="100%"  border="0" cellpadding="3" cellspacing="1" bgcolor="dddddd">
    <tr bgcolor="ffffff"> 
      <td width="100" height="26"> 
        <div align="center">插件名称</div></td>
      <td> 
        <input name="Name" type="text" id="Name" style="width:90%" value="<%=PlusObj("Name")%>"></td>
    </tr>
    <tr bgcolor="ffffff"> 
      <td height="26"> 
        <div align="center">链接地址</div></td>
      <td> 
        <input name="Link" type="text" id="Link" style="width:90%" value="<%=PlusObj("Link")%>"></td>
    </tr>
    <tr bgcolor="ffffff"> 
      <td height="26"> 
        <div align="center">打开方式</div></td>
      <td> 
        <input name="OpenType" type="radio" value="1" <%If PlusObj("OpenType")="1" then Response.Write("checked") end if%>>
        新窗口 
        <input type="radio" name="OpenType" value="0" <%If PlusObj("OpenType")="0" then Response.Write("checked") end if%>>
        原窗口</td>
    </tr>
    <tr bgcolor="ffffff"> 
      <td height="26"> 
        <div align="center">是否显示</div></td>
      <td> 
        <input name="ShowTF" type="radio" value="1" <%If PlusObj("ShowTF")="1" then Response.Write("checked") end if%>>
        显&nbsp;&nbsp;示 
        <input type="radio" name="ShowTF" value="0" <%If PlusObj("ShowTF")="0" then Response.Write("checked") end if%>>
        隐&nbsp;&nbsp;藏</td>
    </tr>
</table>
</form>
</body>
</html>
<%
   if Request.Form("action") = "mod" then
      Dim PlusAddObj,PlusAddSql
	  if NoCSSHackAdmin(Request.Form("Name"),"插件名称") = "" or isnull(Request.Form("Name")) then
	  	 Response.Write("<script>alert(""请输入插件名称"");</script>")
		 Response.End
	  end if
	  if Request.Form("Link") = "" or isnull(Request.Form("Link")) then
	  	 Response.Write("<script>alert(""请输入插件链接地址"");</script>")
		 Response.End
	  end if
	  Set PlusAddObj = Server.CreateObject(G_FS_RS)
	  	  PlusAddSql = "Select * from FS_Plus where ID="&PlusID&""
		  PlusAddObj.Open PlusAddSql,Conn,1,3
		  PlusAddObj("Name") = Replace(Replace(Request.Form("Name"),"""",""),"'","")
		  PlusAddObj("Link") = Request.Form("Link")
		  if Request.Form("OpenType") = "1" then
			  PlusAddObj("OpenType") = "1"
		  else
			  PlusAddObj("OpenType") = "0"
		  end if
		  if Request.Form("ShowTF") = "1" then
			  PlusAddObj("ShowTF") = "1"
		  else
			  PlusAddObj("ShowTF") = "0"
		  end if
		  PlusAddObj.Update
		  PlusAddObj.Close
		  Set PlusAddObj = Nothing
		  Response.Redirect("PlusList.asp")
		  Response.End
   end if
Conn.Close
Set Conn = Nothing
%>