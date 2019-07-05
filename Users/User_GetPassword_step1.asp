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
              <TD width="100%" height="37"> <div align="left"> 
                  <form name=LoginForm method="post" action="User_GetPassword_step2.asp" onSubmit="return CheckLogindata()">
                    <%
Dim RsMemObj
Set RsMemObj = Server.CreateObject (G_FS_RS)
RsMemObj.Source="select * from FS_Members where MemName='"& Replace(Replace(Trim(Request.Form("MemName")),"'","''"),Chr(39),"") &"'"
RsMemObj.Open RsMemObj.Source,Conn,1,1
If Not RsMemObj.eof then
	Call main()
Else
	Call Nonerecode()
End if
Sub main()
%>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td>&nbsp;</td>
                      </tr>
                    </table>
                    <div align="center"> 
                      <table width="44%" border="0" align="center" cellpadding="5" cellspacing="1" bordercolor="e6e6e6" bgcolor="#B4B4B4">
                        <tr bgcolor="#F0F0F0"> 
                          <td width="39%"> 
                            <div align="right">你得密码问题：</div></td>
                          <td width="61%"> 
                            <input name="MemName" type="hidden" id="MemName" value="<%=RsMemObj("MemName")%>"> 
                            <%=RsMemObj("PassQuestion")%></td>
                        </tr>
                        <tr bgcolor="#F0F0F0"> 
                          <td> 
                            <div align="right">请输入你的密码答案：</div></td>
                          <td> 
                            <input name="PassAnswer" type="password" id="PassAnswer"></td>
                        </tr>
                        <tr bgcolor="#F0F0F0"> 
                          <td> 
                            <div align="right">你的电子邮件：</div></td>
                          <td> 
                            <input name="email" type="text" id="PassAnswer3"></td>
                        </tr>
                        <tr bgcolor="#F0F0F0"> 
                          <td colspan="2"> 
                            <div align="center"> 
                              <input name="Submit" type="submit" class="Anbut1" value="下一步">
                            </div></td>
                        </tr>
                      </table>
                    </div>
<%
End sub
Sub Nonerecode()
%>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td>&nbsp;</td>
                      </tr>
                    </table>
                    <table width="80%" border="0" align="center" cellpadding="5" cellspacing="1" bordercolor="e6e6e6" bgcolor="e6e6e6">
                      <tr>
                        <td bgcolor="#FFFFFF"> 
<%
Response.Write("未查询到此用户")
%>
                        </td>
                      </tr>
                    </table>
<%
End sub
%>
                    <div align="center"></div>
                  </form>
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
