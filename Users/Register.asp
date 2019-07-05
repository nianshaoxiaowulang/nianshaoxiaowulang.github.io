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
	dim DBC,conn
	set DBC = new databaseclass
	set conn = DBC.openconnection()
	set DBC = nothing
	Dim I,RsConfigObj
	Set RsConfigObj = Conn.Execute("Select SiteName,UserConfer,Copyright from FS_Config")
%>
<HTML><HEAD><TITLE><%=RsConfigObj("SiteName")%> >> 注册会员</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<META content="MSHTML 6.00.2600.0" name=GENERATOR>
<META content=JavaScript name=vs_defaultClientScript>
<LINK href="Css/UserCSS.css" type=text/css 
rel=stylesheet></HEAD>
<BODY leftMargin=0 topMargin=10 MS_POSITIONING="GridLayout">
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
                                
                              <TD width="85%"><span class="f4"><STRONG>欢迎注册新会员</STRONG></span> 
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
                                      <TD> <font color="#FF0000">同意注册协议</font></TD>
                                    </TR>
                                    <TR> 
                                      <TD width=14 height=24><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD>填写会员资料</TD>
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
class=f4>注册新会员</TD>
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
                  <table width="90%" border="0" align="center" cellpadding="5" cellspacing="0">
                    <tr> 
                      <td> <% = RsConfigObj("UserConfer") %> </td>
                    </tr>
                  </table>
                </div></TD>
            </TR>
            <TR> 
              <TD width="100%" height="47" background=""><div align="center"> 
                  <input style="CURSOR: hand" onclick="window.location.href='sRegister.asp'" type="submit" name="Submit3" value="同意协议">
                  　　 
                  <input style="CURSOR: hand" onclick="javascript:history.go(-1);" type="submit" name="Submit22" value="拒绝以上协议">
                </div></TD>
            </TR>
          </TBODY>
        </TABLE></TD></TR></TBODY></TABLE>
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="0">
  <tr> 
    <td> <hr size="1" noshade color="#FF6600"> <div align="center"> 
        <% = RsConfigObj("Copyright") %>
      </div></td>
  </tr>
</table>
</BODY></HTML>
