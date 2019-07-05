<!--#include file="../Inc/Cls_DB.asp" -->
<!--#include file="../Inc/Const.asp" -->
<!--#include file="../Inc/Md5.asp" -->
<!--#include file="../Inc/NoSqlHack.asp" -->
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
if Request.Form("action")="Check" then
	MemName = Replace(Trim(Request.Form("MemName")),"'","''")
	Password = md5(Request.Form("MemPass"),16)

if MemName = "" or  Password = "" then 
	Response.write"<script>alert(""用户名和密码不能为空"");location.href=""javascript:history.back()"";</script>"
	Response.end
end if
Set RsLoginObj = Server.CreateObject(G_FS_RS)
SqlLogin = "Select * from FS_members where MemName='"&MemName&"' and  password='"&Password&"'"
RsLoginObj.Open SqlLogin,Conn,1,1
if Not RsLoginObj.EOF then 
   if RsloginObj("Lock")=true then
	   Response.write"<script>alert(""您已经被锁定，请联系管理员"");location.href=""javascript:history.back()"";</script>"
	   Response.end
   end if
   Response.Cookies("Foosun")("MemName") = MemName
   Response.Cookies("Foosun")("MemPassword") = Password
   Response.Cookies("Foosun")("MemID") = RsLoginObj("ID")
   Response.Cookies("Foosun")("GroupID") = RsLoginObj("GroupID")
   Session("MemName")=MemName
   Session("MemPassword")=Password
   Session("MemID")=RsLoginObj("ID")
   dim LoginTime
   LoginTime = Now()
   conn.execute("Update FS_members set LoginNum=LoginNum+1,Point=Point+1,LastLoginIP='"&Request.ServerVariables("Remote_ADDR")&"',LastLoginTime='"&LoginTime&"' where MemName='"&MemName&"'")'用户登陆一次，积分+1分
   Response.Redirect("UserIndex.asp") 
   Response.End
else
   Response.write"<script>alert(""非法登陆！请检查用户名和密码的正确性"");location.href=""javascript:history.back()"";</script>"
   Response.end
end if
set Conn = Nothing
Set RsLoginObj = Nothing

end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>会员中心</title>
<style type="text/css">
<!--
 BODY   {border: 0; margin: 0; cursor: default; font-family:宋体; font-size:9pt;}
 BUTTON {width:5em}
 TABLE  {font-family:宋体; font-size:9pt}
 P      {text-align:center}
-->
</style>
</head>
<body leftmargin="0" topmargin="0">
<%
MemName = Session("MemName")
PassWord = Session("MemPassword")
MemID = Session("MemID")
set RsMemObj = Server.CreateObject (G_FS_RS)
RsMemObj.Source="select * from FS_Members where MemName='"& MemName &"' and password='"&PassWord&"'"
RsMemObj.Open RsMemObj.Source,Conn,1,1
if not RsMemObj.EOF then
%>
<table width="226" border="0" cellpadding="2" cellspacing="0">
  <tr> 
    <td colspan="4" class="tabbgcolorlileft"><span class="Nred9pt"><font color="#FF0000"><%=MemName%></font></span><font color="#FF0000">：</font>欢迎您！<a href="Main.asp" target="_top">控制面板</a> 
      <a href="Comm/LetOut.asp" target="_top">退出</a> </td>
  </tr>
  <tr> 
    <td width="67"> <div align="right">一般积分：</div></td>
    <td width="36"><%=RsMemObj("Point")%></td>
    <td width="60">登陆次数：</td>
    <td width="47"><%=RsMemObj("LoginNum")%></td>
  </tr>
  <tr> 
    <td width="67"> <div align="right">注册时间：</div></td>
    <td colspan="3"><%=RsMemObj("RegTime")%></td>
  </tr>
  <tr> 
    <td width="67"> <div align="right">登陆时间：</div></td>
    <td colspan="3"><%=RsMemObj("LastLoginTime")%></td>
  </tr>
  <tr> 
    <td width="67"> <div align="right">登陆ＩＰ：</div></td>
    <td colspan="3"><%=RsMemObj("LastLoginIP")%></td>
  </tr>
  <tr> 
    <td colspan="4"><div align="center"><a href="User_Modify_Pass.asp" target="_top">修改密码</a>　　<a href="User_Modify_contact.asp" target="_top">修改资料</a>　</div></td>
  </tr>
</table>
<%
else
%>
<table width="100%" border="0" cellspacing="5" cellpadding="0">
  <form action="" method="post" name="LoginForm">
    <tr> 
      <td> 
        <div align="center">用户： 
          <input name="MemName" type="text" id="MemName" size="15">
      </div></td>
    </tr>
    <tr> 
      <td> 
        <div align="center">密码： 
          <input name="MemPass" type="password" id="MemPass" size="15">
      </div></td>
    </tr>
    <tr> 
      <td><div align="center">
          <input type="submit" name="Submit" value="登录">
          <input type="reset" name="Submit2" value="重置">
          <input name="action" type="hidden" id="action" value="Check">&nbsp;&nbsp;
          <a href="Register.asp" target="_top"><font color="#FF0000">注册</font></a> 
          <a href="User_GetPassword.asp" target="_top">忘记密码</a></div></td>
    </tr>
  </form>
</table>
  <%
end if
%>
</body>
</html>
