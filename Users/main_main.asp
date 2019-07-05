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
	Dim DBC,conn,sConn
	Set DBC = new databaseclass
	Set Conn = DBC.openconnection()
	Dim I,RsConfigObj
	Set RsConfigObj = Conn.Execute("Select SiteName,UserConfer,Copyright,isEmail,isChange,IsShop from FS_Config")
	If cint(RsConfigObj("IsShop"))=1 Then
		Dim MallConfigObj
		Set MallConfigObj = Conn.execute("select MiddleNum,GoldNum,VipNum from FS_Shop_Config")
	End If
	Set DBC = Nothing
%>
<!--#include file="Comm/User_Purview.Asp" -->
<%
	Dim RsUserObj
	Set RsUserObj = Conn.Execute("Select * From FS_Members where MemName = '"& Replace(session("MemName"),"'","")&"' and Password = '"& Replace(session("MemPassword"),"'","") &"'")
%>
<HTML><HEAD>
<TITLE><%=RsConfigObj("SiteName")%> >> 会员中心</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="Css/UserCSS.css" type=text/css  rel=stylesheet>
</HEAD>
<BODY leftmargin="0" topmargin="10">
<div align="center"> </div>
<TABLE cellSpacing=2 width="98%" align=center border=0>
  <TBODY>
    <TR> 
      <TD vAlign=top> <TABLE cellSpacing=0 cellPadding=5 width="98%" align=center 
                  border=0>
          <TBODY>
            <TR> 
              <TD width="100%"> <TABLE width="100%" border=0>
                  <TBODY>
                    <TR> 
                      <TD width=26><IMG 
                              src="images/Favorite.OnArrow.gif" border=0></TD>
                      <TD 
class=f4>会员中心</TD>
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
                <TD width="100%" height="159"> <div align="left"> 
                    <table width="75%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td height="3"></td>
                      </tr>
                    </table>
                    <TABLE style="BORDER-COLLAPSE: collapse" borderColor=#cccccc 
            cellSpacing=0 cellPadding=0 width="100%" border=1>
                      <TBODY>
                        <TR> 
                          <TD vAlign=top> <TABLE class=bgup cellSpacing=0 cellPadding=5 width="100%" 
                  background="" border=0>
                              <TBODY>
                                <TR> 
                                  <TD width="95%" height="68"><div align="left"><span class="f4"><font color="#FF0000"><strong><%=Session("MemName")%></strong></font></span> <font color="#000000">欢迎您！　　　
								  <%
								  Dim NewsSql,GetMessageObj,TotleMessage
								  NewsSql = "Select * from FS_Message Where MeRead='"& session("memname")&"' and ReadTF=0 and isDelR=0 and IsRecyle=0"
								  Set GetMessageObj = Server.CreateObject(G_FS_RS)
								  GetMessageObj.Open NewsSql,Conn,1,1
								  TotleMessage = GetMessageObj.Recordcount
								  If TotleMessage=0 then
								  	Response.Write("<a href=User_Message.asp>短消息(0)</a>")
								  Else
								  	Response.Write("<a href=User_Message.asp>您有新短消息(<font color=red>"&TotleMessage&"</font>)</a>")
								  End If
								  %>
								  <br>
                                      </font> </div>
                                    <span class="f41"> </span> <table width="75%" border="0" cellspacing="0" cellpadding="5">
                                      <tr> 
                                        <td><span class="f41">用户编号：<font color="#FF0000"> 
                                          <% =  RsUserObj("UserNo") %>
                                          </font> </span></td>
                                      </tr>
                                      <tr> 
                                        <td><span class="f41">注册时间： 
                                          <% =  RsUserObj("RegTime") %>
                                          </span></td>
                                      </tr>
                                      <tr> 
                                        <td>登陆次数：<span class="f41"> 
                                          <% =  RsUserObj("LoginNum") %>
                                          </span></td>
                                      </tr>
                                      <tr> 
                                        <td>一般积分：<span class="f41"> 
                                          <% =  RsUserObj("Point") %>
                                          </span></td>
                                      </tr>
                                      <tr> 
                                        <td><%
										If cint(RsConfigObj("isShop"))=1 then
										%> <hr size="1" noshade color="#FF6600"> 
                                          <table width="100%" border="0" cellpadding="4" cellspacing="1" bgcolor="#DFDFDF">
                                            <tr bgcolor="#FFFFFF"> 
                                              <td height="32" colspan="2"><span class="f41">可用金币： 
                                                <% =  RsUserObj("UserPoint") %>
                                                </span></td>
                                            </tr>
                                            <tr bgcolor="#FFFFFF"> 
                                              <td height="32" colspan="2"> 消费积分：<span class="f41"> 
                                                <% =  RsUserObj("ShopPoint") %>
                                                　　级别： 
                                                <%
												If RsUserObj("ShopPoint")< MallConfigObj("MiddleNum") Then
													Response.Write("<b><font color=""#666666"">一般会员</font></b>")
												Elseif RsUserObj("ShopPoint")>= MallConfigObj("MiddleNum") and RsUserObj("ShopPoint")< MallConfigObj("GoldNum") Then
													Response.Write("<b><font color=""#009900"">中级会员</font></b>")
												Elseif RsUserObj("ShopPoint")>= MallConfigObj("GoldNum") and RsUserObj("ShopPoint")< MallConfigObj("VipNum") Then
													Response.Write("<b><font color=""#0033CC"">高级会员</font></b>")
												Elseif RsUserObj("ShopPoint")>= MallConfigObj("VipNum") Then
													Response.Write("<b><font color=""#990066"">VIP会员</font></b>")
												End if
												%>
                                                </span> </td>
                                            </tr>
                                            <tr bgcolor="#EFEFEF"> 
                                              <td height="25" colspan="2">积分标准：</td>
                                            </tr>
                                            <tr bgcolor="#FFFFFF"> 
                                              <td width="20%"><div align="right"><strong><font color="#666666">一般会员：</font></strong></div></td>
                                              <td width="80%">小于<font color="#FF0000"> 
                                                <% = MallConfigObj("MiddleNum")%>
                                                </font>分</td>
                                            </tr>
                                            <tr bgcolor="#FFFFFF"> 
                                              <td><div align="right"><strong><font color="#009900">中级会员：</font></strong></div></td>
                                              <td>大于或者等于<font color="#FF0000"> 
                                                <% = MallConfigObj("MiddleNum")%>
                                                </font>分，小于<font color="#FF0000"> 
                                                <% = MallConfigObj("GoldNum")%>
                                                </font> 分 </td>
                                            </tr>
                                            <tr bgcolor="#FFFFFF"> 
                                              <td><div align="right"><strong><font color="#0033CC">高级会员：</font></strong></div></td>
                                              <td>大于或者等于 <font color="#FF0000"> 
                                                <% = MallConfigObj("GoldNum")%>
                                                </font>分，小于<font color="#FF0000"> 
                                                <% = MallConfigObj("VIPNum")%>
                                                </font> 分 </td>
                                            </tr>
                                            <tr bgcolor="#FFFFFF"> 
                                              <td><div align="right"><strong><font color="#990066">VIP会员：</font></strong></div></td>
                                              <td>大于或者等于 <font color="#FF0000"> 
                                                <% = MallConfigObj("VIPNum")%>
                                                </font>分</td>
                                            </tr>
                                          </table>
                                          <%
										  Set MallConfigObj = nothing
										  End If
										  %></td>
                                      </tr>
                                    </table></TD>
                                </TR>
                              </TBODY>
                            </TABLE></TD>
                        </TR>
                      </TBODY>
                    </TABLE>
                    <strong></strong></div></TD>
              </form>
            </TR>
          </TBODY>
        </TABLE></TD>
    </TR>
  </TBODY>
</TABLE>
  
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="0">
  <tr>
    <td> 
      <div align="center">
        <hr size="1" noshade color="#FF6600">
        <% = RsConfigObj("Copyright") %>
      </div></td>
  </tr>
</table>
</BODY></HTML>
<%
RsConfigObj.Close
Set RsConfigObj = Nothing
RsUserObj.close
Set RsUserObj=nothing
Set Conn=nothing
%>