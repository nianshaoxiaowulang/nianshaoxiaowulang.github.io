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
	Set Conn = DBC.openconnection()
	Dim I,RsConfigObj
	Set RsConfigObj = Conn.Execute("Select SiteName,UserConfer,Copyright,isEmail,isChange,IsShop from FS_Config")
	Set DBC = Nothing
%>
<!--#include file="Comm/User_Purview.Asp" -->
<%
	Dim RsUserObj
	Set RsUserObj = Conn.Execute("Select Point,RegTime,UserNo,UserPoint,ShopPoint From FS_Members where MemName = '"& Session("MemName")&"' and Password = '"& Session("MemPassword") &"'")
%>
<HTML><HEAD>
<TITLE><%=RsConfigObj("SiteName")%> >> 会员中心</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="Css/UserCSS.css" type=text/css  rel=stylesheet>
</HEAD>
<BODY bgcolor="#FFFFFF" leftmargin="8" topmargin="0">
<div align="center"> </div>
<TABLE width="15%" border=0 align="center" cellpadding="0" cellSpacing=0 bgcolor="#FFFFFF">
  <TBODY>
    <TR> 
      <TD vAlign=top width=160> <a href="main_main.asp" target="main"><IMG src="images/favorite.left.help.jpg" alt="返回管理中心首页" 
        width=190 height=27 border="0"></a>
<TABLE cellSpacing=0 cellPadding=0 width=190 border=0>
          <TBODY>
            <TR> 
              <TD> <TABLE style="BORDER-COLLAPSE: collapse" borderColor=#cccccc 
            cellSpacing=0 cellPadding=0 width="100%" border=1>
                  <TBODY>
                    <TR> 
                      <TD vAlign=top> <TABLE class=bgup cellSpacing=0 cellPadding=0 width="100%" 
                  background="" border=0>
                          <TBODY>
                            <TR> 
                              <TD width="88%" height="30" bgcolor="#FFFFFF"> <table width="75%" border="0" cellspacing="0" cellpadding="0">
                                  <tr> 
                                    <td height="1"></td>
                                  </tr>
                                </table>
                                <div align="center"> </div>
                                <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
                                  <tr> 
                                    <td height="27" bgcolor="#E8E8E8"> <div align="center">基本信息</div></td>
                                  </tr>
                                </table>
                                <TABLE cellSpacing=0 cellPadding=3 width="100%" 
border=0>
                                  <TBODY>
                                    <TR> 
                                      <TD height=13><A><IMG 
                              src="images/arr2.gif" width=10 height=10 id=KB1Img></A></TD>
                                      <TD><a href="All_User.Asp" target="main">注册用户统计</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD height=13 valign="top"><A><IMG 
                              src="images/arr2.gif" width=10 height=10 id=KB1Img></A></TD>
                                      <TD><a href="main_main.asp" target="main">积分与等级</a></TD>
                                    </TR>
                                    <%
								If cint(RsConfigObj("isShop"))=1 then
								%>
                                    <TR> 
                                      <TD height=13 valign="top"><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="Mall/News.asp" target="main">会员公告</a></TD>
                                    </TR>
                                    <%
								End if
								%>
                                    <TR> 
                                      <TD height=6 colspan="2" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                                          <tr> 
                                            <td height="27" bgcolor="#E8E8E8"> 
                                              <div align="center">帐号信息</div></td>
                                          </tr>
                                        </table></TD>
                                    </TR>
                                    <TR> 
                                      <TD width=14 height=13 valign="top"><A><IMG 
                              src="images/arr2.gif" width=10 height=10 id=KB1Img></A></TD>
                                      <TD> <a href="User_Modify_account.asp" target="main">帐号信息</a>/<a href="User_Modify_contact.asp" target="main">联系方式</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD height=13><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="User_Modify_Pass.asp" target="main">修改密码提示答案</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD width=14 height=13><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="User_Modify_other.asp" target="main">其他联系方式</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD height=6 colspan="2" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                                          <tr> 
                                            <td height="27" bgcolor="#E8E8E8"> 
                                              <div align="center">信息管理</div></td>
                                          </tr>
                                        </table></TD>
                                    </TR>
                                    <TR> 
                                      <TD height=13><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="User_Comments.asp" target="main">我发表的评论</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD height=13><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="User_Favorite.asp" target="main">我收藏的信息</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD height=13><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="User_contribution.asp" target="main">稿件管理</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD height=13><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="Add_Contribution.asp" target="main">添加稿件</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD height=13><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="GBook/All_GBook.asp" target="main">留言管理</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD height=13 colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                                          <tr> 
                                            <td height="27" bgcolor="#E8E8E8"> 
                                              <div align="center">短消息</div></td>
                                          </tr>
                                        </table></TD>
                                    </TR>
                                    <TR> 
                                      <TD height=13><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="User_WriteMessage.asp" target="main">撰写消息</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD height=13><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="User_Message.asp?action=Inbox" target="main">收件箱</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD height=13><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="User_Message.asp?action=Outbox" target="main">发件箱</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD height=13><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="User_Message.asp?action=Recycle" target="main">废件箱</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD height=13><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="User_AddressList.asp" target="main">地址薄</a></TD>
                                    </TR>
                                  </TBODY>
                                </TABLE>
                                <%
								If cint(RsConfigObj("isShop"))=1 then
								%> <table width="98%" border="0" cellspacing="0" cellpadding="3">
                                  <tr> 
                                    <td height="27" bgcolor="#E8E8E8"> <div align="center">商城管理</div></td>
                                  </tr>
                                </table>
                                <TABLE cellSpacing=0 cellPadding=3 width="100%" 
border=0>
                                  <TBODY>
                                    <TR> 
                                      <TD width=14 height=13><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD width="154"> <a href="Mall/BuyOrder.asp" target="main">订单管理</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD width=14 height=13><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="Mall/Integral.asp" target="main">我的积分/金币</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD width=14 height=13><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="Mall/Favorite.asp" target="main">收藏夹</a></TD>
                                    </TR>
                                    <TR style="display:"> 
                                      <TD height=13><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="Mall/BuyProductPack.asp" target="main">购物车</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD height=13><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="Mall/Exchange.asp" target="main">积分换金币</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD height="7"><IMG height=5 src="images/SelfService.aspx" 
                              width=1><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="Mall/RegCompany.asp" target="main">注册我的企业</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD height="7"><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="Mall/RegCompanyManage.asp" target="main">修改我的企业</a></TD>
                                    </TR>
                                    <TR>
                                      <TD height="7"><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><A href="mall/Pmf.asp" target="main">配送须知</A></TD>
                                    </TR>
                                  </TBODY>
                                </TABLE>
                                <%
								Else
								%> <table width="98%" border="0" align="center" cellpadding="3" cellspacing="0">
                                  <tr> 
                                    <td height="27" bgcolor="#E8E8E8"> <div align="center">商城管理</div></td>
                                  </tr>
                                </table>
                                <TABLE width="100%" height="27" 
border=0 cellPadding=3 cellSpacing=0>
                                  <TBODY>
                                    <TR> 
                                      <TD height="21"><font color="#FF0000">未开通</font></TD>
                                    </TR>
                                  </TBODY>
                                </TABLE>
                                <%
								End if
								%> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                  <tr> 
                                    <td height="27" bgcolor="#E8E8E8"> <div align="center"><a href="main_main.asp" target="main"><font color="#FF0000">会员中心首页</font></a>｜<a href="Comm/LetOut.asp" target="_top"><font color="#990000">安全退出</font></a></div></td>
                                  </tr>
                                </table></TD>
                            </TR>
                          </TBODY>
                        </TABLE></TD>
                    </TR>
                  </TBODY>
                </TABLE></TD>
            </TR>
          </TBODY>
        </TABLE></TD>
    </TR>
  </TBODY>
</TABLE>
  
<table width="100%" border="0" cellspacing="3" cellpadding="0">
  <tr> 
    <td height="5">风讯官方站：<a href="http://www.Foosun.Cn" target="_blank">Foosun.Cn</a> 
    </td>
  </tr>
  <tr> 
    <td height="5">风讯帮助站：<a href="http://Help.Foosun.Net" target="_blank">Help.Foosun.Net</a></td>
  </tr>
  <tr> 
    <td height="5">风讯交流站：<a href="http://BBS.Foosun.Net" target="_blank">BBS.Foosun.Net</a></td>
  </tr>
</table>
</BODY></HTML>
<%
Sub SendEmail()

End Sub
Sub EmailInfo()
	Response.Write("一封信已经发送到你注册的电子邮件<font color=red>"& Session("email") &"</font>，请注意查收！")
End Sub
RsConfigObj.Close
Set RsConfigObj = Nothing
RsUserObj.close
Set RsUserObj=nothing
Set Conn=nothing
%>