<% Option Explicit %>
<!--#include file="../Inc/Cls_DB.asp" -->
<!--#include file="../Inc/Const.asp" -->
<!--#include file="../Inc/Md5.asp" -->
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
Set RsConfigObj = Conn.Execute("Select * from FS_Config")
'发送邮件
Function SendMail(SMTPServer,sender,loginName,LoginPass,mailto,subject,mailbody)'这个是一个函数，在程序中可以直接调用。
	On error resume next
	Dim JMail
	Set JMail = Server.CreateObject("JMail.Message") 
	JMail.Silent = True
	JMail.Logging = True
	JMail.Charset = "gb2312" '设定邮件的字符集，默认为"US-ASCII" 一般中文用"gb2312" 
	If Not(LoginName = "" Or LoginPass = "") Then
		JMail.MailServerUserName = LoginName '您的邮件服务器登录名
		JMail.MailServerPassword = LoginPass '登录密码
	End If
	JMail.ContentType = "text/html" '这个邮件的头设置， 默认为 "text/plain" 能设置成你需要的任何情况。 '你想发送HTML信息,改这个头文件为 "text/html"
	JMail.From = sender'邮寄的地址
	JMail.FromName = ""&RsConfigObj("SiteName")&"网站管理员"
	JMail.Subject = subject'邮件的标题。 
	JMail.AddRecipient mailto'加入一个收件者
	JMail.Body = mailbody 'E-Mail的主体
	JMail.Priority = 1'邮件的优先级，可以范围从1到5。越大的优先级约高，比如，5最高，1最低,一般设置
	JMail.AddHeader "Originating-IP", Request.ServerVariables("REMOTE_ADDR")'addheader加入一个头信息
	if not JMail.Send(SMTPServer) then'执行邮件发送到SERVER  
		SendMail = false
		Response.Write("邮件发送失败，可能是服务器不支持JMAIL组件，请使用jmail4.3以上版本！<br>")
	Else
		SendMain = true
		Response.Write("邮件已经发送到你注册的邮箱中，请注意查收<br>")
	End If
	JMail.Close
	Set JMail=nothing   
End Function
%>
<HTML><HEAD>
<TITLE><%=RsConfigObj("SiteName")%> >> 注册成功</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="Css/UserCSS.css" type=text/css  rel=stylesheet>
<meta http-equiv="refresh" content="10;URL=main.asp">
</HEAD>
<BODY leftmargin="0" topmargin="10">
<div align="center">
  <script language="JavaScript" src="top.js" type="text/JavaScript"></script>
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
                                      <TD>填写帐号信息</TD>
                                    </TR>
                                    <TR> 
                                      <TD width=14 height=24><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD>填写联系资料</TD>
                                    </TR>
                                    <TR> 
                                      <TD width=14 height=24><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><font color="#FF0000">注册成功</font></TD>
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
class=f4>注册成功</TD>
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
              <form method=POST action="sRegister_Success.asp" name=UserForm1 onSubmit="return checkdata1()">
                <TD width="100%" height="159"> 
                  <div align="center"><br>
                    恭喜您!<font color="#FF0000"><strong><%=Session("sName")%></strong></font>，您在<%=RsConfigObj("SiteName")%>注册成功<br>
                    <br>
					<%
					if cint(RsConfigObj("isEmail"))=1 then
						Call SendMail(trim(RsConfigObj("MailServer")),trim(RsConfigObj("Email")),trim(RsConfigObj("MailName")),RsConfigObj("MailPass"),Session("email"),"来自["&RsConfigObj("SiteName")&"]的注册信息",Session("sName")&":您好！<br>欢迎注册成为"&RsConfigObj("SiteName")&"会员,您的用户名："& Session("MemName") &"，密码："& Session("VerPassword") &",登陆地址："&RsConfigObj("DoMain")&"/User/Login.asp")
						Response.Write("<br>")
					End if
					%>
                    10秒后返回<a href="main.asp"><font color="#FF0000">会员中心主页</font></a> 
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
<%
RsConfigObj.Close
Set RsConfigObj = Nothing
Set Conn=nothing
%>