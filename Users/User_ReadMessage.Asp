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
'如需进行2次开发，必须经过风讯公司书面允许。否则将追究法律责任
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
	Dim RsUserObj
	Set RsUserObj = Conn.Execute("Select * From FS_Members where MemName = '"& Replace(Replace(session("MemName"),"'",""),Chr(39),"")&"' and Password = '"& Replace(Replace(session("MemPassword"),"'",""),Chr(39),"") &"'")
	If RsUserObj.eof then
		Response.Write("<script>alert(""严重错误！"&CopyRight&""");location=""Login.asp"";</script>")  
		Response.End
	End if
	If Request.Form("action")="ReSend" Then
		If Request.Form("MeTitle")="" Or Request.Form("MeFrom")="" Or Request.Form("MeContent")="" Then
			Response.Write("<script>alert(""请填写完整！"&CopyRight&""");location=""javascript:history.back()"";</script>")  
			Response.End
		End If
		If len(Request.Form("MeContent"))>300 Then
			Response.Write("<script>alert(""短信内容不能超过300个字节！"&CopyRight&""");location=""javascript:history.back()"";</script>")  
			Response.End
		End If
		If Trim(Request.Form("MeFrom"))=session("MemName") Then
			Response.Write("<script>alert(""不能给自己发送短信！"&CopyRight&""");location=""javascript:history.back()"";</script>")  
			Response.End
		End if
		Dim GetTFobj
		Set GetTFobj=Conn.execute("select * from FS_members where MemName ='"& replace(Trim(Request.Form("MeFrom")),"'","")&"'")
		If GetTFobj.eof then
			Response.Write("<script>alert(""没有此用户！"&CopyRight&""");location=""javascript:history.back()"";</script>")  
			Response.End
		End if
		Dim SumRsObj,TotleSQL
		Set SumRsObj = Server.CreateObject(G_FS_RS)
		TotleSQL = "Select sum(LenContent) from FS_Message where MeRead='"& replace(Trim(Request.Form("MeFrom")),"'","") &"' and IsDelR = 0"
		SumRsObj.Open TotleSQL,Conn,1,3
		If SumRsObj(0)> 50*1024  Or SumRsObj(0)+Len(Request.Form("MeContent")) > 50*1024 then
			Response.Write("<script>alert(""对方短信空间容量已满！请通知对方删除多余电子邮件"&CopyRight&""");location=""javascript:history.back()"";</script>")  
			Response.End
		End If
		Dim SendRsObj,SendSQL
		Set SendRsObj = Server.CreateObject(G_FS_RS)
		SendSQL = "Select * from FS_Message where 1=0"
		SendRsObj.Open SendSQL,Conn,1,3
		SendRsObj.addnew
		SendRsObj("MeTitle")=Replace(Request.Form("MeTitle"),"'","")
		SendRsObj("MeFrom")=Session("MemName")
		SendRsObj("MeRead")=Replace(Request.Form("MeFrom"),"'","")
		SendRsObj("MeContent")=Request.Form("MeContent")
		SendRsObj("FromDate")=now
		SendRsObj("ReadTF")=0
		SendRsObj("IsRecyle")=0
		SendRsObj("IsDels")=0
		SendRsObj("IsDelR")=0
		SendRsObj("LenContent")=Len(Request.Form("MeContent"))
		SendRsObj.Update
		SendRsObj.Close
		Set SendRsObj=nothing
		Response.Write("<script>alert(""恭喜！\b发送成功！"&CopyRight&""");location=""User_Message.asp"";</script>")  
		Response.End
	End If
	If Request("id")="" Then
		Response.Write("<script>alert(""错误参数！"&CopyRight&""");location=""javascript:history.back()"";</script>")  
		Response.End
	End If
	Dim RsReadMessageObj
	Set RsReadMessageObj = Conn.execute("Select * from FS_Message where MeId="& clng(Replace(Request("id"),"'","")))
	If RsReadMessageObj.Eof Then
		Response.Write("<script>alert(""错误参数！"&CopyRight&""");location=""javascript:history.back()"";</script>")  
		Response.End
	End If
	If RsReadMessageObj("MeRead")<> session("MemName") Then
		Response.Write("<script>alert(""错误参数！"&CopyRight&""");location=""javascript:history.back()"";</script>")  
		Response.End
	End If
	Conn.execute("Update FS_Message Set ReadTF=1 where MeId="& clng(Replace(Request("id"),"'","")))
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
class=f4>读取短消息</TD>
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
            cellSpacing=0 cellPadding=5 width="100%" border=1>
                    <TBODY>
                        <TR> 
                          <TD vAlign=top> <TABLE class=bgup cellSpacing=0 cellPadding=5 width="100%" 
                  background="" border=0>
                            <TBODY>
                              <TR> 
                                <TD width="6%" height="26"> <div align="left"><font color="#000000"> 
                                    </font> <font color="#000000"> </font> </div>
                                  <a href="User_Message.asp?action=Inbox"><img src="Images/o_inbox.gif" width="40" height="40" border="0"></a> 
                                </TD>
                                <TD width="6%"><a href="User_Message.asp?action=Outbox"><img src="Images/M_outbox.gif" width="40" height="40" border="0"></a></TD>
                                <TD width="6%"><a href="User_Message.asp?action=Recycle"><img src="Images/M_recycle.gif" width="40" height="40" border="0"></a></TD>
                                <TD width="6%"><a href="User_AddressList.asp"><img src="Images/M_address.gif" width="40" height="40" border="0"></a></TD>
                                <TD width="2%"><span class="f41"><a href="User_WriteMessage.asp"><img src="Images/m_write.gif" width="40" height="40" border="0"></a></span></TD>
                                <TD width="68%"><div align="center"></div></TD>
                              </TR>
                            </TBODY>
                          </TABLE>
                          <hr size="1" noshade>
                          
                            <table width="100%" height="98" border="0" cellpadding="5" cellspacing="1" bgcolor="#E8E8E8">
                            <form name="form1" method="post" action="">
                              <tr bgcolor="#FFFFFF"> 
                                <td width="19%" height="31"> 
                                  <div align="right">标题： 
                                  </div></td>
                                <td width="81%"> 
                                  <input name="MeTiTle" type="text" id="MeTiTle" value="Re:<% = RsReadMessageObj("MeTiTle")%>" size="50"></td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td> 
                                  <div align="right">发信人：</div></td>
                                <td> 
                                  <input name="MeFrom" type="text" id="MeFrom" value="<% = RsReadMessageObj("MeFrom")%>" size="50">
                                  <font color="#999999">回复只能回复单个用户！</font></td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td> 
                                  <div align="right">内容正文：</div></td>
                                <td> 
<textarea name="MeContent" cols="60" rows="9" id="MeContent">=====<% = RsReadMessageObj("MeFrom")%> 于<% = RsReadMessageObj("FromDate")%>说道：===================

<% = RsReadMessageObj("MeContent")%></textarea> 
                                  <font color="#666666">最多300个字符。 </font></td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td> 
                                  <div align="right">发送时间：</div></td>
                                <td> 
                                  <% = RsReadMessageObj("FromDate")%></td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td>&nbsp;</td>
                                <td> 
                                  <input type="submit" name="Submit" value="回复">
                                  <input name="action" type="hidden" id="action" value="ReSend"></td>
                              </tr>
                            </form>
                          </table>
                           </TD>
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
RsUserObj.close
Set RsUserObj=nothing
Set Conn=nothing
%>

