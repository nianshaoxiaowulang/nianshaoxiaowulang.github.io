<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
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
if Not JudgePopedomTF(Session("Name"),"P040402") then Call ReturnError1()
	Dim UserIDP,UserModifyObj,TempUserName
	UserIDP = Request("ID")
	Set UserModifyObj = Conn.Execute("Select * from FS_Members where ID="&Clng(UserIDP)&"")
	If UserModifyObj.eof then
		Response.Write("<script>alert(""参数传递错误"");window.close();</script>")
		Response.End
	End If
	TempUserName = UserModifyObj("MemName")
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet" type="text/css">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>修改会员信息</title>
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
          <td>&nbsp; <input name="action" type="hidden" id="action" value="mod"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="100%"  border="0" cellpadding="3" cellspacing="1" bgcolor="#DDDDDD">
    <tr> 
      <td width="9%" bgcolor="#EAEAEA">
<div align="right">会&nbsp;员&nbsp;名</div></td>
      <td width="37%" bgcolor="#FFFFFF">
<input name="MemName" readonly type="text" id="MemName" style="width:90%" value="<%=UserModifyObj("MemName")%>"></td>
      <td width="12%" bgcolor="#EAEAEA">
<div align="right">会&nbsp;员&nbsp;组</div></td>
      <td width="39%" bgcolor="#FFFFFF">
<select name="GroupID" id="GroupID" style="width:90%">
          <option value="0" <%If UserModifyObj("GroupID") = "" or  UserModifyObj("GroupID") = "0" then Response.Write("selected") end if%>> 
          </option>
          <%
		Dim SelGroupObj
		Set SelGroupObj = Conn.Execute("Select GroupID,Name from FS_MemGroup order by PopLevel desc")
		do while not SelGroupObj.eof 
	%>
          <option value="<%=SelGroupObj("GroupID")%>" <%If Cstr(UserModifyObj("GroupID"))=Cstr(SelGroupObj("GroupID")) then Response.Write("selected") end if%>><%=SelGroupObj("Name")%></option>
          <%
		SelGroupObj.MoveNext
		Loop
		SelGroupObj.Close
		Set SelGroupObj = Nothing
	%>
        </select></td>
    </tr>
    <tr> 
      <td bgcolor="#EAEAEA">
<div align="right">密&nbsp;&nbsp;&nbsp;&nbsp;码</div></td>
      <td bgcolor="#FFFFFF">
<input name="Password" type="password" id="Password" style="width:90%"></td>
      <td bgcolor="#EAEAEA">
<div align="right">确认密码</div></td>
      <td bgcolor="#FFFFFF">
<input name="PasswordTF" type="password" id="PasswordTF2" style="width:90%"></td>
    </tr>
    <tr> 
      <td bgcolor="#EAEAEA">
<div align="right">真实姓名</div></td>
      <td bgcolor="#FFFFFF">
<input name="Name" type="text" id="Name2" style="width:90%" value="<%=UserModifyObj("Name")%>"></td>
      <td bgcolor="#EAEAEA">
<div align="right">电话号码</div></td>
      <td bgcolor="#FFFFFF">
<input name="Telephone" type="text" id="Telephone2" style="width:90%" value="<%=UserModifyObj("Telephone")%>"></td>
    </tr>
    <tr> 
      <td bgcolor="#EAEAEA">
<div align="right">E_Mail</div></td>
      <td bgcolor="#FFFFFF">
<input name="Email" type="text" id="Email2" style="width:90%" value="<%=UserModifyObj("Email")%>"></td>
      <td bgcolor="#EAEAEA">
<div align="right">OICQ</div></td>
      <td bgcolor="#FFFFFF">
<input name="Oicq" type="text" id="Oicq2" style="width:90%" value="<%=UserModifyObj("Oicq")%>"></td>
    </tr>
    <tr> 
      <td bgcolor="#EAEAEA">
<div align="right">MSN</div></td>
      <td bgcolor="#FFFFFF">
<input name="MSN" type="text" id="MSN2" style="width:90%" value="<%=UserModifyObj("MSN")%>"></td>
      <td bgcolor="#EAEAEA">
<div align="right">个人主页</div></td>
      <td bgcolor="#FFFFFF">
<input name="HomePage" type="text" id="HomePage2" style="width:90%" value="<%=UserModifyObj("HomePage")%>"></td>
    </tr>
    <tr> 
      <td bgcolor="#EAEAEA">
<div align="right">地&nbsp;&nbsp;&nbsp;&nbsp;址</div></td>
      <td bgcolor="#FFFFFF">
<input name="Address" type="text" id="Address2" style="width:90%" value="<%=UserModifyObj("Address")%>"></td>
      <td bgcolor="#EAEAEA">
<div align="right">区域国家</div></td>
      <td bgcolor="#FFFFFF">
<input name="Corner" type="text" id="Corner2" style="width:90%" value="<%=UserModifyObj("Corner")%>"></td>
    </tr>
    <tr> 
      <td bgcolor="#EAEAEA">
<div align="right">省&nbsp;&nbsp;&nbsp;&nbsp;份</div></td>
      <td bgcolor="#FFFFFF">
<input name="Province" type="text" id="Province2" style="width:90%" value="<%=UserModifyObj("Province")%>"></td>
      <td bgcolor="#EAEAEA">
<div align="right">城&nbsp;&nbsp;&nbsp;&nbsp;市</div></td>
      <td bgcolor="#FFFFFF">
<input name="City" type="text" id="City2" style="width:90%" value="<%=UserModifyObj("City")%>"></td>
    </tr>
    <tr> 
      <td bgcolor="#EAEAEA">
<div align="right">职&nbsp;&nbsp;&nbsp;&nbsp;业</div></td>
      <td bgcolor="#FFFFFF">
<input name="Vocation" type="text" id="Vocation2" style="width:90%" value="<%=UserModifyObj("Vocation")%>"></td>
      <td bgcolor="#EAEAEA">
<div align="right">学&nbsp;&nbsp;&nbsp;&nbsp;历</div></td>
      <td bgcolor="#FFFFFF">
<input name="EduLevel" type="text" id="EduLevel2" style="width:90%" value="<%=UserModifyObj("EduLevel")%>"></td>
    </tr>
    <tr> 
      <td bgcolor="#EAEAEA">
<div align="right">登录次数</div></td>
      <td bgcolor="#FFFFFF">
<input name="LoginNum" type="text" id="LoginNum2" style="width:90%" value="<%=UserModifyObj("LoginNum")%>"></td>
      <td bgcolor="#EAEAEA">
<div align="right">积&nbsp;&nbsp;&nbsp;&nbsp;分</div></td>
      <td bgcolor="#FFFFFF">
<input name="Point" type="text" id="Point2" style="width:90%" value="<%=UserModifyObj("Point")%>"></td>
    </tr>
    <tr> 
      <td bgcolor="#EAEAEA">
<div align="right">密码提示</div></td>
      <td bgcolor="#FFFFFF">
<input name="PassQuestion" type="text" id="PassQuestion2" style="width:90%" value="<%=UserModifyObj("PassQuestion")%>"></td>
      <td bgcolor="#EAEAEA">
<div align="right">问题答案</div></td>
      <td bgcolor="#FFFFFF">
<input name="PassAnswer" type="text" id="PassAnswer2" style="width:90%" value="<%=UserModifyObj("PassAnswer")%>"></td>
    </tr>
    <tr> 
      <td bgcolor="#EAEAEA">
<div align="right">介 绍 人</div></td>
      <td bgcolor="#FFFFFF">
<input name="AddUser" type="text" id="AddUser2" style="width:90%" value="<%=UserModifyObj("AddUser")%>"></td>
      <td bgcolor="#EAEAEA">
<div align="right">头像图片</div></td>
      <td bgcolor="#FFFFFF">
<input name="HeadPic" type="text" id="HeadPic2" style="width:90%" value="<%=UserModifyObj("HeadPic")%>"></td>
    </tr>
    <tr> 
      <td bgcolor="#EAEAEA">
<div align="right">签&nbsp;&nbsp;&nbsp;&nbsp;名</div></td>
      <td bgcolor="#FFFFFF">
<input name="UnderWrite" type="text" id="UnderWrite2" style="width:90%" value="<%=UserModifyObj("UnderWrite")%>"></td>
      <td bgcolor="#EAEAEA">
<div align="right">生&nbsp;&nbsp;&nbsp;&nbsp;日</div></td>
      <td bgcolor="#FFFFFF">
<input name="Birthday" type="text" id="Birthday2" style="width:90%" value="<%=UserModifyObj("Birthday")%>"></td>
    </tr>
    <tr> 
      <td bgcolor="#EAEAEA">
<div align="right">最后登录</div></td>
      <td bgcolor="#FFFFFF">
<input name="LastLoginIP" type="text" id="LastLoginIP2" style="width:90%" value="<%=UserModifyObj("LastLoginIP")%>" readonly></td>
      <td bgcolor="#EAEAEA">
<div align="right">登录时间</div></td>
      <td bgcolor="#FFFFFF">
<input name="LastLoginTime" type="text" id="LastLoginTime2" style="width:90%" value="<%=UserModifyObj("LastLoginTime")%>" readonly></td>
    </tr>
    <tr> 
      <td bgcolor="#EAEAEA">
<div align="right">锁&nbsp;&nbsp;&nbsp;&nbsp;定</div></td>
      <td bgcolor="#FFFFFF">
<input type="radio" name="Lock" value="1" <%If UserModifyObj("Lock") = "1" then Response.Write("checked") end if%>>
        是 
        <input name="Lock" type="radio" value="0" <%If UserModifyObj("Lock") = "0" or UserModifyObj("Lock") = "" then Response.Write("checked") end if%>>
        否</td>
      <td bgcolor="#EAEAEA">
<div align="right">性&nbsp;&nbsp;&nbsp;&nbsp;别</div></td>
      <td bgcolor="#FFFFFF">
<input name="Sex" type="radio" value="0" <%If UserModifyObj("Sex") = "0" or UserModifyObj("Sex") = "" then Response.Write("checked") end if%>>
        男 
        <input type="radio" name="Sex" value="1" <%If UserModifyObj("Sex") = "1" then Response.Write("checked") end if%>>
        女</td>
    </tr>
    <tr> 
      <td bgcolor="#EAEAEA">
<div align="right">开放资料</div></td>
      <td bgcolor="#FFFFFF">
<input name="OpenInfTF" type="radio" value="1" <% if UserModifyObj("OpenInfTF")="1" then Response.Write("checked") end if%>>
        是 
        <input type="radio" name="OpenInfTF" value="0" <% if UserModifyObj("OpenInfTF")="0" then Response.Write("checked") end if%>>
        否</td>
      <td bgcolor="#EAEAEA">
<div align="right">订阅信息</div></td>
      <td bgcolor="#FFFFFF">
<input name="SubInfTF" type="radio" value="1" <% if UserModifyObj("SubInfTF")="1" then Response.Write("checked") end if%>>
        是 
        <input type="radio" name="SubInfTF" value="0" <% if UserModifyObj("SubInfTF")="0" then Response.Write("checked") end if%>>
        否</td>
    </tr>
    <tr> 
      <td bgcolor="#EAEAEA">
<div align="right">个人介绍</div></td>
      <td colspan="3" bgcolor="#FFFFFF">
<textarea name="SelfIntro" rows="6" id="SelfIntro" style="width:90%"><%=UserModifyObj("SelfIntro")%></textarea>
      </td>
    </tr>
</table>
</form>
</body>
</html>
<%
If  Request.Form("action") = "mod" then
    Dim UserAddObj,UserAddSql,ChooseMemNameObj,MemNameStr
	If NoCSSHackAdmin(Request.Form("MemName"),"会员名")="" or isnull(Request.Form("MemName")) then
		Response.Write("<script>alert(""请填写会员登录名"");</script>")
		Response.End
	Else
		MemNameStr = Replace(Replace(Request.Form("MemName"),"""",""),"'","")
	End If
	Set ChooseMemNameObj = Conn.Execute("Select count(ID) from FS_Members where MemName='"&MemNameStr&"'")
	If Cstr(UserModifyObj("MemName"))=Cstr(MemNameStr) and ChooseMemNameObj(0)>1 then
		Response.Write("<script>alert(""此会员登录名已经存在,请修改"");</script>")
		Response.End
	Elseif Cstr(UserModifyObj("MemName"))<>Cstr(MemNameStr) and ChooseMemNameObj(0)<>0 then
		Response.Write("<script>alert(""此会员登录名已经存在,请修改"");</script>")
		Response.End
	End If
	ChooseMemNameObj.Close
	Set ChooseMemNameObj = Nothing
	If Request.Form("Password")<>"" and Cstr(Request.Form("Password"))<>Cstr(Request.Form("PasswordTF")) then
		Response.Write("<script>alert(""密码与确认密码不同"");</script>")
		Response.End
	End If
	If Request.Form("Name")="" or isnull(Request.Form("Name")) then
		Response.Write("<script>alert(""请填写会员真实姓名"");</script>")
		Response.End
	End If
	'===========================================
	'判断输入的生日格式，正确才保存
	If Request.Form("Birthday")<>"" then
		If Not Isdate(Request.Form("Birthday")) then 
		Response.Write("<script>alert(""请输入正确的生日格式,例：1976-1-1"");</script>")
		Response.End
		end if
	End If
	'==========================================
	Set UserAddObj = Server.CreateObject(G_FS_RS)
		UserAddSql = "Select * from FS_Members where ID="&UserIDP&""
		UserAddObj.Open UserAddSql,Conn,3,3
		UserAddObj("MemName") = Replace(Replace(Request.Form("MemName"),"""",""),"'","")
		If Request.Form("Password")<>"" then
			UserAddObj("Password") = md5(Request.Form("Password"),16)
		End If
		UserAddObj("GroupID") = Request.Form("GroupID")
		UserAddObj("Name") = Replace(Replace(Request.Form("Name"),"""",""),"'","")
		UserAddObj("Email") = Request.Form("Email")
		UserAddObj("Telephone") = Request.Form("Telephone")
		UserAddObj("Oicq") = Request.Form("Oicq")
		UserAddObj("HomePage") = Request.Form("HomePage")
		UserAddObj("LoginNum") = Request.Form("LoginNum")
		UserAddObj("Address") = Request.Form("Address")
		UserAddObj("SelfIntro") = Request.Form("SelfIntro")
		UserAddObj("AddUser") = Request.Form("AddUser")
		if Request.Form("Point") <> "" then
			UserAddObj("Point") = Request.Form("Point")
		end if
		UserAddObj("HeadPic") = Request.Form("HeadPic")
		If Request.Form("Birthday")<>"" then
			UserAddObj("Birthday") = Request.Form("Birthday")
		End If
		UserAddObj("MSN") = Request.Form("MSN")
		UserAddObj("Corner") = Request.Form("Corner")
		UserAddObj("Province") = Request.Form("Province")
		UserAddObj("City") = Request.Form("City")
		UserAddObj("Vocation") = Request.Form("Vocation")
		UserAddObj("EduLevel") = Request.Form("EduLevel")
		If Request.Form("OpenInfTF") = "0" then
			UserAddObj("OpenInfTF") = "0"
		Else
			UserAddObj("OpenInfTF") = "1"
		End If
		If Request.Form("SubInfTF") = "0" then
			UserAddObj("SubInfTF") = "0"
		Else
			UserAddObj("SubInfTF") = "1"
		End If
		UserAddObj("PassQuestion") = Request.Form("PassQuestion")
		UserAddObj("PassAnswer") = Request.Form("PassAnswer")
		UserAddObj("UnderWrite") = Request.Form("UnderWrite")
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
		UserAddObj.Update
		UserAddObj.Close
		Set UserAddObj = Nothing
		If Cstr(TempUserName) <> Cstr(Replace(Replace(Request.Form("MemName"),"""",""),"'","")) then
			Conn.Execute("Update FS_Message set MeRead='"&Replace(Replace(Request.Form("MemName"),"""",""),"'","")&"' where MeRead='"&TempUserName&"'")
		End If
		Response.Redirect("SysUserList.asp")
		Response.End
End If
%>