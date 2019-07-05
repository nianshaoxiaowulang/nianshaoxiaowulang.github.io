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
If Request("StrPage1")<>"" then
	Response.Redirect("All_User.Asp?page="& Request("StrPage1") &"&Keyword="& Request("Keyword") &"")
	Response.end
End if
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
class=f4>所有注册用户</TD>
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
                    
                  <table width="100%" border="0" cellpadding="6" cellspacing="1" bgcolor="#DDDDDD">
                    <tr bgcolor="#F3F3F3"> 
                      <td colspan="3">本站共注册 
                        <%
						Dim RsCountObj,SQL
						Dim strpage
						strpage=request.querystring("page")
						if len(strpage)=0 then
							strpage="1"
						end if
						Set RsCountObj = Server.CreateObject(G_FS_RS)
						If Request("Keyword")<>"" then
							SQL = "Select * From Fs_Members where MemName Like '%" & Replace(Replace(Request("Keyword"),"'",""),Chr(39),"") & "%' Order by ID desc"
						Else
							SQL = "Select * From Fs_Members Order by ID desc"
						End if
						RsCountObj.Open SQL,Conn,1,1
							Response.Write "<Font color=red>" & RsCountObj.RecordCount&"</font>"
						%>
                        个会员 </td>
                       <form name="form1" method="post" action="All_User.Asp">
					  <%
					  	If cint(RsConfigObj("IsShop"))=1 Then
					  %>
					   <td colspan="4"> 
					   <%Else%>
					   <td colspan="2"> 
					   <%End if%><div align="center">
                          <input name="Keyword" type="text" id="Keyword" value="<%=Request("Keyword")%>">
                          <input type="submit" name="Submit" value="搜索">
                        </div></td>
                     
                      </form>
                    </tr>
                    <tr bgcolor="#EAEAEA"> 
                      <td width="15%"><strong>用户名</strong></td>
                      <td width="13%"><strong>腾讯QQ</strong></td>
                      <td width="16%"><strong>注册时间</strong></td>
                      <td width="18%"><strong>最后登陆时间</strong></td>
					  <%
					  	If cint(RsConfigObj("IsShop"))=1 Then
					  %>
                      <td width="13%"><strong>等级状态</strong></td>
					  <%
					  End if
					  %>
                      <td width="14%"><strong>登陆积分</strong></td>
					  <%
					  	If cint(RsConfigObj("IsShop"))=1 Then
					  %>
                      <td width="11%"><strong>消费积分</strong></td>
					  <%
					  End if
					  %>
                    </tr>
                    <%
					    Dim select_count,select_pagecount
						if RsCountObj.eof then
							   RsCountObj.close
							   set RsCountObj=nothing
							   Response.Write"<TR><TD colspan=""9""  bgcolor=""#FFFFFF"">没有记录。</TD></TR>"
						else
								RsCountObj.pagesize = 20
								RsCountObj.absolutepage=cint(strpage)
								select_count=RsCountObj.recordcount
								select_pagecount=RsCountObj.pagecount
								for i=1 to RsCountObj.pagesize
									if RsCountObj.eof then
										exit for
									end if
					  %>
                    <tr bgcolor="#FFFFFF"> 
                      <td><A href="ReadUser.Asp?UserName=<% = RsCountObj("MemName")%>" target="_blank"> 
                        <% = RsCountObj("MemName")%>
                        </a></td>
                      <td> <%
						if Trim(RsCountObj("OICQ"))<>"" then
							Dim sOICQ
						    sOICQ ="<a target=blank href=http://wpa.qq.com/msgrd?V=1&Uin="& RsCountObj("OICQ") &"&Site="& RsConfigObj("SiteName") &"&Menu=yes><img border=""0"" SRC=http://wpa.qq.com/pa?p=1:"& RsCountObj("OICQ") &":8 alt=""点击这里给"& RsCountObj("OICQ") &"发消息""></a>"
							Response.Write sOICQ
						Else
							Response.Write("没有")
						End if
						%></td>
                      <td> <% = RsCountObj("RegTime")%></td>
                      <td>
                        <% = RsCountObj("LastLoginTime")%>
                      </td>
					  <%
					  	If cint(RsConfigObj("IsShop"))=1 Then
					  %>
                      <td> <%
							If RsCountObj("ShopPoint")< MallConfigObj("MiddleNum") Then
								Response.Write("<b><font color=""#666666"">一般会员</font></b>")
							Elseif RsCountObj("ShopPoint")>= MallConfigObj("MiddleNum") and RsCountObj("ShopPoint")< MallConfigObj("GoldNum") Then
								Response.Write("<b><font color=""#009900"">中级会员</font></b>")
							Elseif RsCountObj("ShopPoint")>= MallConfigObj("GoldNum") and RsCountObj("ShopPoint")< MallConfigObj("VipNum") Then
								Response.Write("<b><font color=""#0033CC"">高级会员</font></b>")
							Elseif RsCountObj("ShopPoint")>= MallConfigObj("VipNum") Then
								Response.Write("<b><font color=""#990066"">VIP会员</font></b>")
							End if
							%></td>
						<%
						End if
						%>
                      <td> <% = RsCountObj("Point")%></td>
					  <%
					  	If cint(RsConfigObj("IsShop"))=1 Then
					  %>
                      <td> <% = RsCountObj("ShopPoint")%></td>
					  <%End if%>
                    </tr>
                    <%
						  RsCountObj.MoveNext
					  Next
					  %>
                  </table>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="52%">
                        <% 
							Response.write"<br>&nbsp;共<b>"& select_pagecount &"</b>页<b>&nbsp;" & select_count &"</b>条记录，本页是第<b>"& strpage &"</b>页。"
							if int(strpage)>1 then
								Response.Write"&nbsp;<a href=?page=1&Keyword="&Request("Keyword")&">第一页</a>&nbsp;&nbsp;"
								Response.Write"&nbsp;<a href=?page="&cstr(cint(strpage)-1)&"&Keyword="&Request("Keyword")&">上一页</a>&nbsp;&nbsp;"
							End if
							If int(strpage)<select_pagecount then
								Response.Write"&nbsp;<a href=?page="&cstr(cint(strpage)+1)&"&Keyword="&Request("Keyword")&">下一页</a>&nbsp;"
								Response.Write"&nbsp;<a href=?page="& select_pagecount &"&Keyword="&Request("Keyword")&">最后一页</a>&nbsp;&nbsp;"
							End if
								Response.Write"<br>"
								RsCountObj.close
								Set RsCountObj=nothing
							End if
							%>
                      </td>
                      <form name="form2" method="post" action=""><td width="48%">
                          转到第几页： 
                          <input name="StrPage1" type="text" id="StrPage1" value="<%=Request("Page")%>" size="3">
                          <input type="submit" name="Submit2" value="转到">
                        </td></form>
                    </tr>
                  </table>
                </div></TD>
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