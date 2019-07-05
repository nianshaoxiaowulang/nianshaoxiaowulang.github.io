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
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P040204") then Call ReturnError()
Dim AdminID,Result
Result = Request("Result")
AdminID = Replace(Replace(Request("AdminID"),"'",""),"""","")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>修改管理员密码</title>
</head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet">
<body>
<div align="center">
  <table width="95%" border="0" cellspacing="0" cellpadding="0">
  <form action="" method="post" name="PassWordForm">
    <tr>
      <td height="15" colspan="2">&nbsp;</td>
      </tr>
    <tr> 
      <td width="23%" height="30"> 
        <div align="left">&nbsp;&nbsp;&nbsp;新 密 码</div></td>
      <td width="77%" height="30"> 
        <div align="left"> 
          <input name="PassWord" type="password" id="PassWord" style="width:100%;">
        </div></td>
    </tr>
    <tr> 
      <td height="30"><div align="left">&nbsp;&nbsp;&nbsp;确认密码</div></td>
      <td height="30"><div align="left"><input name="AffirmPassWord" type="password" id="AffirmPassWord" style="width:100%;">
        </div></td>
    </tr>
    <tr> 
      <td height="50" colspan="2"> 
        <div align="center"> 
                  <input type="submit" name="Submit" value=" 确 定 ">
                    <input name="AdminID" value="<% = AdminID %>" type="hidden" id="AdminID">
                    <input name="Result" type="hidden" id="Result" value="Submit">
                    <input type="button" onClick="window.close();" name="Submit2" value=" 取 消 ">
        </div></td>
    </tr>
 </form>
  </table>
</div>
</body>
</html>
<%
if Result = "Submit" then
	Dim PassWord,AffirmPassWord,RsAdminObj,ReturnCheckInfo
	PassWord = Replace(Replace(Request.Form("PassWord"),"'",""),"""","")
	AffirmPassWord = Replace(Replace(Request.Form("AffirmPassWord"),"'",""),"""","")
	AdminID = Replace(Replace(Request.Form("AdminID"),"'",""),"""","")
	Set RsAdminObj = Server.CreateObject(G_FS_RS)
	RsAdminObj.Open "Select * from FS_Admin where ID="& AdminID &"",Conn
	if RsAdminObj.Eof then
		Set Conn = Nothing
		%>
		<script>alert('此管理员已经被删除');dialogArguments.location.reload();window.close();</script>
		<%
	end if
	if Len(PassWord) < 6 then
		Set Conn = Nothing
		%>
		<script>alert('密码至少要六位');</script>
		<%
		Response.End  
	end if
	if PassWord <> AffirmPassWord then
		Set Conn = Nothing
		%>
		<script>alert('确认密码不对');</script>
		<%
		Response.End  
	end if
	On Error Resume Next
	Conn.Execute("update FS_Admin set PassWord='" & md5(PassWord,16) & "' where ID=" & AdminID & "")
	if Err.Number = 0 then
		%>
		<script>dialogArguments.location.reload();window.close();</script>
		<%
	else
		%>
		<script>alert('发生错误');</script>
		<%
	end if
end if
Set Conn = Nothing
%>