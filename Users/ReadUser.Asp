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
Set DBC = Nothing
Dim RsUserObj,UserSQL
Set RsUserObj = Server.CreateObject(G_FS_RS)
UserSQL = "select * from FS_Members where MemName='"& Replace(Replace(Request("UserName"),"'",""),Chr(39),"") &"'"
RsUserObj.Open UserSQL,Conn,1,3
If RsUserObj.eof Then
	Response.Write("<script>alert(""找不到用户,或者已经被删除！"&CopyRight&""");location=""javascript:history.back()"";</script>")  
	Response.End
Else
%>
<HTML><HEAD>
<TITLE><%=RsConfigObj("SiteName")%> >> 会员中心</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="Css/UserCSS.css" type=text/css  rel=stylesheet>
</HEAD>
<BODY leftmargin="0" topmargin="5">
<div align="center">
  <script language="JavaScript" src="top.js" type="text/JavaScript"></script>
</div>
<TABLE width="98%" border=0 align=center cellpadding="0" cellSpacing=0>
  <TBODY>
    <TR> 
      <TD vAlign=top> <TABLE cellSpacing=0 cellPadding=0 width="100%" align=center 
                  border=0>
          <TBODY>
            <TR> 
              <TD width="100%"> <TABLE width="100%" border=0 cellpadding="0" cellspacing="0">
                  <TBODY>
                    <TR> 
                      <TD width=26><IMG 
                              src="images/Favorite.OnArrow.gif" border=0></TD>
                      <TD 
class=f4>用户资料</TD>
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
                <TD width="100%" height="159" valign="top"> 
                  <div align="left"> 
                    <table width="75%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td height="3"></td>
                      </tr>
                    </table>
                    
                  <TABLE style="BORDER-COLLAPSE: collapse" borderColor=#cccccc 
            cellSpacing=0 cellPadding=0 width="100%" border=1>
                    <TBODY>
                        <TR> 
                          
                        <TD height="207" vAlign=top><table width="100%" border="0" cellpadding="8" cellspacing="1" bordercolor="#CCCCCC" bgcolor="#E6E6E6">
                            <tr bgcolor="#EEEEEE"> 
                              <td colspan="4"><font color="#FF3300"><strong>基本信息</strong></font></td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td><div align="right">用户编号：</div></td>
                              <td> <% = RsUserObj("UserNo")%> </td>
                              <td width="13%" rowspan="2"><div align="right">头像：</div>
                                <div align="right"></div>
                                <div align="right"></div></td>
                              <td width="39%" rowspan="2"> <%
							  If RsUserObj("HeadPic")<>"" Then
							  %> <table width="0" border="0" cellpadding="0" cellspacing="0" bgcolor="#F0F0F0">
                                  <tr> 
                                    <td bgcolor="#FFFFFF"><img src="<% = RsUserObj("HeadPic")%>" width="40" height="40" border="0"></td>
                                  </tr>
                                </table>
                                <%
							  Else
							  %> <table width="0" border="0" cellpadding="0" cellspacing="0" bgcolor="#F0F0F0">
                                  <tr> 
                                    <td bgcolor="#FFFFFF"><img src="images/noHeadPic.jpg" width="40" height="40" border="0"></td>
                                  </tr>
                                </table>
                                <%
							  End If
							  %></td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td width="17%"><div align="right">用户名：</div></td>
                              <td width="31%"><% = RsUserObj("MemName")%></td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td><div align="right">性别：</div></td>
                              <td> <%
								If RsUserObj("sex")=0 Then
									Response.Write("男")
								Else
									Response.Write("女")
								End If
								%> </td>
                              <td width="13%"><div align="right">最后登陆时间：</div></td>
                              <td width="39%"> <% = RsUserObj("LastLoginTime")%> </td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td><div align="right">省份：</div></td>
                              <td> <% = RsUserObj("Province")%> </td>
                              <td><div align="right">城市：</div>
                                <div align="right"></div>
                                <div align="right"></div></td>
                              <td> <% = RsUserObj("City")%> </td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td><div align="right">注册时间：</div></td>
                              <td><% = RsUserObj("RegTime")%></td>
                              <td><div align="right">登陆次数：</div></td>
                              <td> <% = RsUserObj("LoginNum")%> </td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td><div align="right">消费积分：</div></td>
                              <td> <% = RsUserObj("ShopPoint")%> </td>
                              <td><div align="right">其他积分：</div></td>
                              <td> <% = RsUserObj("Point")%> </td>
                            </tr>
                            <tr bgcolor="#F5F5F5"> 
                              <td height="96"> <div align="right">介绍：</div></td>
                              <td> <% = RsUserObj("SelfIntro")%> </td>
                              <td> <div align="right">签名<strong>：</strong></div></td>
                              <td> <% = RsUserObj("UnderWrite")%> </td>
                            </tr>
                            <tr bgcolor="#EEEEEE"> 
                              <td colspan="4"> <div align="left"><font color="#FF3300"><strong>联系方式</strong></font></div></td>
                            </tr>
                            <%
							If RsUserObj("OpenInfTF") = 0 then
							%>
                            <tr bgcolor="#FFFFFF"> 
                              <td height="38" colspan="4"><font color="#0000FF">用户已经设置为保密</font></td>
                            </tr>
                            <%
							Else
							%>
                            <tr bgcolor="#FFFFFF"> 
                              <td><div align="right">姓名：</div></td>
                              <td colspan="3"> <% = RsUserObj("Name")%> </td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td><div align="right">电话：</div></td>
                              <td colspan="3"> <% = RsUserObj("Telephone")%> </td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td><div align="right">地址：</div></td>
                              <td colspan="3"> <% = RsUserObj("address")%> </td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td><div align="right">邮政编码：</div></td>
                              <td colspan="3"> <% = RsUserObj("PostCode")%> </td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td><div align="right">职业：</div></td>
                              <td colspan="3"> <% = RsUserObj("Vocation")%> </td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td><div align="right">学历：</div></td>
                              <td colspan="3"> <% = RsUserObj("EduLevel")%> </td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td><div align="right">出生日期：</div></td>
                              <td colspan="3"> <% = RsUserObj("birthday")%> </td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td><div align="right">网站地址：</div></td>
                              <td colspan="3"> <% = RsUserObj("Homepage")%> </td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td><div align="right">QQ：</div></td>
                              <td colspan="3"><%
						if Len(RsUserObj("OICQ"))=0 Or Trim(RsUserObj("OICQ"))="" then
							Response.Write("没有")
						Else
							Dim sOICQ
						    sOICQ ="<a target=blank href=http://wpa.qq.com/msgrd?V=1&Uin="& RsUserObj("OICQ") &"&Site="& RsConfigObj("SiteName") &"&Menu=yes><img border=""0"" SRC=http://wpa.qq.com/pa?p=1:"& RsUserObj("OICQ") &":8 alt=""点击这里给"& RsUserObj("OICQ") &"发消息""></a>"
							Response.Write sOICQ
						End if
						%> </td>
                            </tr> 
                            <tr bgcolor="#FFFFFF"> 
                              <td><div align="right">MSN：</div></td>
                              <td colspan="3"> <% = RsUserObj("msn")%> </td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td><div align="right">Email:</div></td>
                              <td colspan="3"> <a href="mailto:<% = RsUserObj("email")%>"><% = RsUserObj("email")%></a></td>
                            </tr>
                            <%End If%>
                          </table> </TD>
                        </TR>
                      </TBODY>
                    </TABLE>
                    <strong></strong></div></TD>
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
Set Conn=nothing
End If
%>

