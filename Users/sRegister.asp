<% Option Explicit %>
<!--#include file="../Inc/Cls_DB.asp" -->
<!--#include file="../Inc/Const.asp" -->
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
	Set RsConfigObj = Conn.Execute("Select SiteName,UserConfer,Copyright from FS_Config")
%>
<HTML><HEAD><TITLE><%=RsConfigObj("SiteName")%> >> 填写帐号信息</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="Css/UserCSS.css" type=text/css  rel=stylesheet></HEAD>
<BODY leftmargin="0" topmargin="10">
<div align="center">
  <script language="JavaScript" src="top.js" type="text/JavaScript"></script>
  <script language="javascript" src="Comm/MyScript.js"></script>
</div>
<TABLE cellSpacing=2 width="98%" align=center border=0>
  <TBODY>
  <TR>
    <TD vAlign=top width=160>
      <TABLE cellSpacing=0 cellPadding=0 width=102 border=0>
        <TBODY>
        <TR>
          <TD><IMG height=27 src="images/favorite.left.help.jpg" 
        width=190></TD></TR>
        <TR>
          <TD>
            <TABLE style="BORDER-COLLAPSE: collapse" borderColor=#cccccc 
            cellSpacing=0 cellPadding=0 width="100%" border=1>
              <TBODY>
              <TR>
                <TD vAlign=top>
                  <TABLE class=bgup cellSpacing=0 cellPadding=0 width="100%" 
                  background="" border=0>
                    <TBODY>
                    <TR>
                      <TD align=right>&nbsp;</TD>
                      <TD align=right>&nbsp;</TD></TR>
                    <TR>
                      <TD align=right width="15%" height=30>&nbsp;</TD>
                                
                              <TD width="85%"><span class="f4"><STRONG>填写帐号信息</STRONG></span> 
                              </TD>
                              </TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
            <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
              <TBODY>
              <TR>
                <TD height=10><IMG height=1 src="" 
width=1></TD></TR></TBODY></TABLE>
            <TABLE style="BORDER-COLLAPSE: collapse" borderColor=#cccccc 
            cellSpacing=0 cellPadding=0 width="100%" border=1>
              <TBODY>
              <TR>
                <TD vAlign=top>
                  <TABLE class=bgup cellSpacing=0 cellPadding=0 width="100%" 
                  background="" border=0>
                    <TBODY>
                    <TR>
                      <TD align=right>&nbsp;</TD>
                      <TD align=right>&nbsp;</TD></TR>
                    <TR>
                      <TD align=right width="12%" height=30>&nbsp;</TD>
                      <TD width="88%">
                        <TABLE cellSpacing=0 cellPadding=0 width="100%" 
border=0>
                                  <TBODY>
                                    <TR> 
                                      <TD width=14 height=24><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD> 同意注册协议</TD>
                                    </TR>
                                    <TR> 
                                      <TD width=14 height=24><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><font color="#FF0000">填写帐号信息</font></TD>
                                    </TR>
                                    <TR> 
                                      <TD width=14 height=24><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD>填写联系资料</TD>
                                    </TR>
                                    <TR> 
                                      <TD width=14 height=24><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD>注册成功</TD>
                                    </TR>
                                    <TR> 
                                      <TD><IMG height=5 src="images/SelfService.aspx" 
                              width=1></TD>
                                      <TD></TD>
                                    </TR>
                                  </TBODY>
                                </TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD>
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
class=f4><span class="f4">填写帐号信息</span></TD>
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
              <form method=POST action="sRegister_1.asp" name=UserForm onSubmit="return checkdata()">
                <TD width="100%" height="74"> <div align="left"> <br>
                    <table width="98%" border="0" cellspacing="0" cellpadding="5">
                      <tr> 
                        <td width="19%"><div align="right">用户名：</div></td>
                        <td width="29%"><input name="Username" type="text" id="Username">
                          <font color="#FF0000"> * <a href="javascript:CheckName('Comm/CheckName.asp')"><font color="#999900"><br>
                          检查是否被占用</font></a> </font></td>
                        <td width="52%">登录时使用的代号，请记住该客户名，建议使用email地址，以便记忆！ </td>
                      </tr>
                      <tr> 
                        <td><div align="right">密码：</div></td>
                        <td><input name="sPassword" type="password" id="sPassword"> 
                          <font color="#FF0000">*</font></td>
                        <td>为了您的个人资料的安全，请您最好使用6位以上密码！ </td>
                      </tr>
                      <tr> 
                        <td><div align="right">验证密码：</div></td>
                        <td><input name="Confimpass" type="password" id="Confimpass"> 
                          <font color="#FF0000">*</font></td>
                        <td>必须和上面输入的密码一致！ </td>
                      </tr>
                      <tr> 
                        <td><div align="right">电子邮件：</div></td>
                        <td><input name="email" type="text" id="email"> <font color="#FF0000">*<br>
                          <a href="javascript:CheckEmail('Comm/Checkemail.asp')"><font color="#999900">检查是否被占用</font></a></font></td>
                        <td>请输入您最常使用的电子邮件地址（每个Email只能注册一位用户）。 </td>
                      </tr>
                      <tr> 
                        <td><div align="right">密码提示问题：</div></td>
                        <td><input name="PassQuestion" type="text" id="PassQuestion">
                          <font color="#FF0000">*</font></td>
                        <td>密码提示问题，用于提示您找回密码的根据</td>
                      </tr>
                      <tr> 
                        <td><div align="right">密码提示答案：</div></td>
                        <td><input name="PassAnswer" type="text" id="PassAnswer">
                          <font color="#FF0000">*</font></td>
                        <td>密码提示答案，用于提示您找回密码的唯一根据之一</td>
                      </tr>
                      <tr> 
                        <td>&nbsp;</td>
                        <td colspan="2"><input  type=submit name="Submit3" value="下一步" style="cursor:hand;"></td>
                      </tr>
                      <tr> 
                        <td>&nbsp;</td>
                        <td colspan="2">注意：<br>
                          1．带*的栏目必须填写，否则注册不能继续！ <br>
                          2．推荐您使用网易2G超大免费邮箱@126.com，点击 <a href="http://reg.126.com/reg1.jsp" target="_blank"><font color="#FF0000">快速注册</font></a> 
                        </td>
                      </tr>
                    </table>
                  </div></TD></form>
            </TR>
          </TBODY>
        </TABLE></TD></TR></TBODY></TABLE>
  
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="0">
  <tr>
    <td> <hr size="1" noshade color="#FF6600">
      <div align="center">
        <% = RsConfigObj("Copyright") %>
      </div></td>
  </tr>
</table>
<BR>
</BODY></HTML>
<script language="JavaScript" type="text/JavaScript">
function CheckName(gotoURL) {
   var ssn=UserForm.Username.value.toLowerCase();
	   var open_url = gotoURL + "?Username=" + ssn;
	   window.open(open_url,'','status=0,directories=0,resizable=0,toolbar=0,location=0,scrollbars=0,width=150,height=80');
}
function CheckEmail(gotoURL) {
   var ssn1=UserForm.email.value.toLowerCase();
	   var open_url = gotoURL + "?email=" + ssn1;
	   window.open(open_url,'','status=0,directories=0,resizable=0,toolbar=0,location=0,scrollbars=0,width=150,height=80');
}
</script>
