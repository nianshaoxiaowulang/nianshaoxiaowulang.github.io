<% Option Explicit %>
<!--#include file="../Inc/Cls_DB.asp" -->
<!--#include file="../Inc/Const.asp" -->
<!--#include file="../Inc/Md5.asp" -->
<!--#include file="../Inc/Function.asp" -->
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
'==============================================================================
	Dim DBC,conn
	Set DBC = new databaseclass
	Set conn = DBC.openconnection()
	Set DBC = nothing
	Dim I,RsConfigObj
	Set RsConfigObj = Conn.Execute("Select SiteName,Copyright from FS_Config")
If request.form("action")="modify" then
	If  Replace(Trim(Request.Form("password")),"'","''") =""  then 
			Response.Write("<script>alert(""密码不能为空！"");location=""Javascript:history.go(-1)"";</script>")
			Response.End
	End if
	If  Replace(Trim(Request.Form("password")),"'","''") <> Replace(Trim(Request.Form("password1")),"'","''")  then
			Response.Write("<script>alert(""2次密码不相同，请重新输入！"");location=""Javascript:history.go(-1)"";</script>")
			Response.End
	End if
	Set RsMemObj = Server.CreateObject (G_FS_RS)
	RsMemObj.Source="select * from FS_Members where ID="& CLNG(Replace(Request.Form("MemID"),Chr(39),""))
	RsMemObj.Open RsMemObj.Source,Conn,1,3
	RsMemObj("Password") = md5(Replace(Trim(Request.Form("password")),"'","''"),16)
	RsMemObj.update
	Response.Write("<script>alert(""修改密码成功！"");location=""Login.asp"";</script>")  
	Response.End
End if
%>
<HTML><HEAD><TITLE><%=RsConfigObj("SiteName")%> >> 会员登陆</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<META content="MSHTML 6.00.2600.0" name=GENERATOR>
<META content=JavaScript name=vs_defaultClientScript>
<LINK href="Css/UserCSS.css" type=text/css rel=stylesheet>
</HEAD>
<script language="javascript" src="Comm/MyScript.js"></script>
<BODY leftMargin=0 topMargin=10 MS_POSITIONING="GridLayout">
<div align="center">
  <script language="JavaScript" src="top.js" type="text/JavaScript"></script>
</div>
<TABLE cellSpacing=2 width="98%" align=center border=0>
  <TBODY>
    <TR> 
      <TD vAlign=top> <TABLE cellSpacing=0 cellPadding=0 width="98%" align=center 
                  border=0>
          <TBODY>
            <TR> 
              <TD width="100%"> <TABLE width="100%" border=0>
                  <TBODY>
                    <TR> 
                      <TD width=26><IMG 
                              src="images/Favorite.OnArrow.gif" border=0></TD>
                      <TD 
class=f4>会员登陆</TD>
                    </TR>
                  </TBODY>
                </TABLE></TD>
            </TR>
            <TR> 
              <TD width="100%"> <TABLE cellSpacing=0 cellPadding=0 width="100%" 
border=0>
                  <TBODY>
                    <TR> 
                      <TD bgColor=#ff6633 height=4><IMG height=1 src="" 
                              width=1></TD>
                    </TR>
                  </TBODY>
                </TABLE></TD>
            </TR>
            <TR> 
              <TD width="100%" height="37"> <div align="left"> <br>
<%
Dim RsMemObj
set RsMemObj = Server.CreateObject (G_FS_RS)
RsMemObj.Source="select * from FS_Members where MemName='"& Replace(Trim(Request.Form("MemName")),"'","''") &"' and PassAnswer='"& MD5(Replace(Trim(Request.Form("PassAnswer")),"'","''"),16) &"'and email='"& Replace(Trim(Request.Form("email")),"'","''") &"'"
RsMemObj.Open RsMemObj.Source,Conn,1,1
If Not RsMemObj.eof then
	Call main()
Else
	Call Nonerecode()
End if
Sub main()%>
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td>&nbsp;</td>
                    </tr>
                  </table>
                  <form name="form1" method="post" action="">
                    <table width="44%" border="0" align="center" cellpadding="5" cellspacing="1" bordercolor="e6e6e6" bgcolor="#B6B6B6">
                      <tr bgcolor="#EEEEEE"> 
                        <td width="23%"> 
                          <div align="right">修改密码：</div></td>
                        <td width="77%"><font color="#FF0000"> 
                          <input name="action" type="hidden" id="action" value="modify">
                          <input name="MemID" type="hidden" value="<%=RsMemObj("ID")%>">
                          <input name="Password" type="password" id="Password">
                          </font><a href="UserLogin.asp"><font color="#0000FF"> 
                          </font></a></td>
                      </tr>
                      <tr bgcolor="#EEEEEE"> 
                        <td> 
                          <div align="right">确认密码：</div></td>
                        <td><font color="#FF0000"> 
                          <input name="Password1" type="password" id="Password1">
                          </font></td>
                      </tr>
                      <tr bgcolor="#EEEEEE"> 
                        <td>&nbsp;</td>
                        <td><a href="UserLogin.asp"><font color="#0000FF"> 
                          <input type="submit" name="Submit" value="修改">
                          </font></a></td>
                      </tr>
                    </table>
                  </form>
<%
End sub
Sub Nonerecode()
%>
                  <table width="80%" border="0" align="center" cellpadding="5" cellspacing="1" bordercolor="e6e6e6" bgcolor="e6e6e6">
                    <tr>
                      <td bgcolor="#FFFFFF"> 
                        <%
						Response.Write("你输入资料不正确，请重新输入！")
						%>
                      </td>
                    </tr>
                  </table>
<%
End sub
%>
                </div></TD>
            </TR>
          </TBODY>
        </TABLE></TD>
    </TR>
  </TBODY>
</TABLE>
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="0">
  <tr> 
    <td> <hr size="1" noshade color="#FF6600"> <div align="center"> 
        <% = RsConfigObj("Copyright") %>
      </div></td>
  </tr>
</table>
</BODY></HTML>
