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
if Not JudgePopedomTF(Session("Name"),"P080502") then Call ReturnError1()
	Dim RsOCObj,TempFlag
	Set RsOCObj = Conn.Execute("Select * from FS_WebInfo")
	If RsOCObj.eof then
		TempFlag = false
	Else
		TempFlag = true
	End If
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>网站维护</title>
</head>
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<body topmargin="2" leftmargin="2" oncontextmenu="//return false;">
<form action="" method="post" name="VOForm">
  <table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr> 
      <td height="28" class="ButtonListLeft"> <div align="center"><strong>网站信息维护</strong></div></td>
    </tr>
  </table>
  <br>
  <table width="75%"  border="0" align="center" cellpadding="3" cellspacing="1" bordercolor="e6e6e6" bgcolor="dddddd">
    <tr bgcolor="#FFFFFF"> 
      <td width="24%">&nbsp;&nbsp;&nbsp;&nbsp;网站名称</td>
      <td width="76%"> 
        <input name="WebName" type="text" id="WebName" style="width:90%" value="<%If TempFlag = true then Response.Write(RsOCObj("WebName")) end if%>"></td>
  </tr>
    <tr bgcolor="#FFFFFF"> 
      <td>&nbsp;&nbsp;&nbsp;&nbsp;网站地址</td>
      <td> 
        <input name="WebUrl" type="text" id="WebUrl" style="width:90%" value="<%If TempFlag = true then Response.Write(RsOCObj("WebUrl")) end if%>"></td>
  </tr>
    <tr bgcolor="#FFFFFF"> 
      <td>&nbsp;&nbsp;&nbsp;&nbsp;管理员</td>
      <td> 
        <input name="WebAdmin" type="text" id="WebAdmin" style="width:90%" value="<%If TempFlag = true then Response.Write(RsOCObj("WebAdmin")) end if%>"></td>
  </tr>
    <tr bgcolor="#FFFFFF"> 
      <td>&nbsp;&nbsp;&nbsp;&nbsp;网站信箱</td>
      <td> 
        <input name="WebEmail" type="text" id="WebEmail" style="width:90%" value="<%If TempFlag = true then Response.Write(RsOCObj("WebEmail")) end if%>"></td>
  </tr>
    <tr bgcolor="#FFFFFF"> 
      <td>&nbsp;&nbsp;&nbsp;&nbsp;开始统计时间</td>
      <td> 
        <input name="WebCountTime" type="text" readonly id="WebCountTime" style="width:71%" value="<%If TempFlag = true then Response.Write(RsOCObj("WebCountTime")) end if%>">
      <input type="button" name="dfgdf" value="选择日期" onClick="OpenWindowAndSetValue('../../FunPages/SelectDate.asp',280,110,window,document.VOForm.WebCountTime);"></td>
  </tr>
    <tr bgcolor="#FFFFFF"> 
      <td>&nbsp;&nbsp;&nbsp;&nbsp;网站介绍</td>
      <td> 
        <textarea name="WebIntro" id="WebIntro" style="width:90%"><%If TempFlag = true then Response.Write(RsOCObj("WebIntro")) end if%></textarea></td>
  </tr>
    <tr bgcolor="#FFFFFF"> 
      <td colspan="2"> 
        <div align="center">
      <input type="submit" name="Submit" value=" 确 定 ">&nbsp;&nbsp;
      <input name="action" type="hidden" id="action" value="trues">
      <input type="reset" name="Submit" value=" 还 原 ">&nbsp;&nbsp;
      <input type="button" name="Submit" value=" 取 消 " onclick="history.back();">
    </div></td>
    </tr>
</table>
</form>
</body>
</html>
<%
	If Request.Form("action") = "trues" then
		Dim VOModObj,VoModSql
		Set VOModObj = Server.CreateObject(G_FS_RS)
		VoModSql = "Select * from FS_WebInfo order by ID asc"
		VOModObj.Open VoModSql,Conn,3,3
		If TempFlag = false then
		VOModObj.AddNew
		End If
		VOModObj("WebName") = Replace(Replace(Request.Form("WebName"),"""",""),"'","")
		VOModObj("WebUrl") = Request.Form("WebUrl")
		VOModObj("WebIntro") = Request.Form("WebIntro")
		VOModObj("WebEmail") = Request.Form("WebEmail")
		VOModObj("WebAdmin") = Request.Form("WebAdmin")
		VOModObj("WebCountTime") = Request.Form("WebCountTime")
		VOModObj.Update
		VOModObj.Close
		Set VOModObj = Nothing
		Response.Write("<script>alert(""网站信息维护成功"");history.back();</script>")
	End If
Conn.Close
Set Conn = Nothing
%>