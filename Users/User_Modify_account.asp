<% Option Explicit %>
<!--#include file="../Inc/Function.asp" -->
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
%>
<!--#include file="Comm/User_Purview.Asp" -->
<%
	If Request.Form("action")="Update" then
		Conn.execute("Update FS_Members Set VerGetType='"&NoCSSHackInput(Replace(Replace(Request.Form("VerGetType"),"'",""),Chr(39),""))&"' where id = "&Clng(Replace(Replace(Request.Form("id"),"'",""),Chr(39),"")))
		If Trim(Request.Form("VerGetCode"))<>"" then
			Conn.execute("Update FS_Members Set VerGetCode='"&md5(Replace(Request.Form("VerGetCode"),"'",""),32)&"' where id = "&Clng(Replace(Replace(Request.Form("id"),"'",""),Chr(39),"")))
		End if
		Response.Write("<script>alert(""更新成功！"&CopyRight&""");location=""User_Modify_account.asp"";</script>")  
		Response.End
	End if
	Dim RsUserObj
	Set RsUserObj = Conn.Execute("Select * From FS_Members where MemName = '"& Replace(Replace(session("MemName"),"'",""),Chr(39),"")&"' and Password = '"& Replace(Replace(session("MemPassword"),"'",""),Chr(39),"") &"'")
	If RsUserObj.eof then
		Response.Write("<script>alert(""严重错误！"&CopyRight&""");location=""Login.asp"";</script>")  
		Response.End
	End if
%>
<HTML><HEAD>
<TITLE><%=RsConfigObj("SiteName")%> >> 会员中心</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="Css/UserCSS.css" type=text/css  rel=stylesheet>
</HEAD>
<BODY leftmargin="0" topmargin="5">
<div align="center"> </div>
<TABLE cellSpacing=2 width="98%" align=center border=0>
  <TBODY>
    <TR> 
      <TD vAlign=top> <TABLE cellSpacing=0 cellPadding=0 width="98%" align=center 
                  border=0>
          <TBODY>
            <TR> 
              <TD width="100%"> <TABLE width="100%" border=0 cellpadding="0" cellspacing="0">
                  <TBODY>
                    <TR> 
                      <TD width=26><IMG 
                              src="images/Favorite.OnArrow.gif" border=0></TD>
                      <TD 
class=f4>修改帐号</TD>
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
              <form method=POST action="" name=UserForm1">
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
                                  <TD width="95%" height="68"><div align="left"><font color="#000000"> 
                                      </font> 
                                      <table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#E7E7E7">
                                        <tr bgcolor="#FFFFFF"> 
                                          <td><div align="right">会员编号：</div></td>
                                          <td><font color="#FF0000"> 
                                            <% = RsUserObj("UserNo") %>
                                            &nbsp;</font></td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td width="22%"> <div align="right">用户名：</div></td>
                                          <td width="78%"> <% = RsUserObj("MemName") %> </td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td> <div align="right">电子邮件：</div></td>
                                          <td> <% = RsUserObj("Email") %> <font color="#666666">，电子邮件不能修改，如果您要修改，请与管理员联系 
                                            </font></td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td> <div align="right">证件类型：</div></td>
                                          <td><font color="#FF0000"> 
                                            <select name="VerGetType" id="VerGetType">
                                              <option value="身份证" <%if RsUserObj("VerGetType")="身份证" then response.Write("selected")%>>身份证</option>
                                              <option value="学生证" <%if RsUserObj("VerGetType")="学生证" then response.Write("selected")%>>学生证</option>
                                              <option value="军人证" <%if RsUserObj("VerGetType")="军人证" then response.Write("selected")%>>军人证</option>
                                              <option value="护照" <%if RsUserObj("VerGetType")="护照" then response.Write("selected")%>>护照</option>
                                            </select>
                                            </font></td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td> <div align="right">身份证：</div></td>
                                          <td> <input name="VerGetCode" type="text" id="VerGetCode"> 
                                            <font color="#666666">不修改，请保持为空 </font></td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td> <div align="right">注册时间：</div></td>
                                          <td> <% = RsUserObj("RegTime") %> </td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td><div align="right">登陆次数：</div></td>
                                          <td> <% = RsUserObj("Point") %> </td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td><div align="right">消费积分：</div></td>
                                          <td> <% = RsUserObj("ShopPoint") %> </td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td><div align="right">其他积分：</div></td>
                                          <td> <% = RsUserObj("Point") %> </td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td><div align="right"><span class="f41">可用金币</span>：</div></td>
                                          <td> <% = RsUserObj("UserPoint") %> </td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td><div align="right">最后登陆时间：</div></td>
                                          <td> <% = RsUserObj("LastLoginTime") %> </td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td><div align="right">最后登陆IP：</div></td>
                                          <td> <% = RsUserObj("LastLoginIP") %> </td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td><div align="right">投稿数量：</div></td>
                                          <td> <% = RsUserObj("ConNum") %> </td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td><div align="right">状态：</div></td>
                                          <td> <% 
											If  RsUserObj("Lock") = 1 Then
												Response.Write("锁定")
											Else
												Response.Write("开放")
											End if
											%> </td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td>&nbsp;</td>
                                          <td> <input type="submit" name="Submit" value="修改身份证及身份证类型"> 
                                            <input name="id" type="hidden" id="id" value="<% = RsUserObj("ID") %>"> 
                                            <input name="action" type="hidden" id="action" value="Update"></td>
                                        </tr>
                                      </table>
                                      <font color="#000000"> </font> </div>
                                    <span class="f41"> </span> </TD>
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