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
%>
<!--#include file="Comm/User_Purview.Asp" -->
<%
If request.Form("action")="delInbox" Then
	If trim(Request.Form("InboxID"))<>"" Then
		Conn.execute("Update FS_Message Set IsDelR=0,IsRecyle=1 Where MeId in("&Request.Form("InboxID")&")")
		Response.Write("<script>alert(""短信删除到回收站成功！"&CopyRight&""");location=""User_Message.asp"";</script>")  
		Response.End
	Else
		Response.Write("<script>alert(""请选择删除的短信！"&CopyRight&""");location=""User_Message.asp"";</script>")  
		Response.End
	End if
End If
If request.Form("action")="RecycleInbox" Then
	If trim(Request.Form("RecycleboxID"))<>"" Then
		Conn.execute("Update FS_Message Set IsDelR=1,IsRecyle=1 Where MeId in("&Request.Form("RecycleboxID")&")")
		Response.Write("<script>alert(""彻底删除成功！"&CopyRight&""");location=""User_Message.asp?action=Recycle"";</script>")  
		Response.End
	Else
		Response.Write("<script>alert(""请选择删除的短信！"&CopyRight&""");location=""User_Message.asp?action=Recycle"";</script>")  
		Response.End
	End if
End If
If request.Form("action")="OutBox" Then
	If trim(Request.Form("OutBoxID"))<>"" Then
		Conn.execute("Update FS_Message Set IsDelR=1,IsRecyle=1,isSend=0 Where MeId in("&Request.Form("OutBoxID")&")")
		Response.Write("<script>alert(""彻底删除成功！-"&CopyRight&""");location=""User_Message.asp?action=Outbox"";</script>")  
		Response.End
	Else
		Response.Write("<script>alert(""请选择删除的短信！"&CopyRight&""");location=""User_Message.asp?action=Outbox"";</script>")  
		Response.End
	End if
End If
Dim RsUserObj
Set RsUserObj = Conn.Execute("Select * From FS_Members where MemName = '"& Replace(session("MemName"),"'","")&"' and Password = '"& Replace(session("MemPassword"),"'","") &"'")
If RsUserObj.eof then
	Response.Write("<script>alert(""严重错误！"&CopyRight&""");location=""Login.asp"";</script>")  
    Response.End
Else
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
	class=f4>短消息</TD>
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
							  
							<TD height="446" vAlign=top> 
							  <TABLE class=bgup cellSpacing=0 cellPadding=5 width="100%" 
					  background="" border=0>
								<TBODY>
								  <TR> 
									<TD width="6%" height="26"> <div align="left"><font color="#000000"> 
										</font> <font color="#000000"> </font> </div>
									  <a href="?action=Inbox"><img src="Images/o_inbox.gif" width="40" height="40" border="0"></a> 
									</TD>
									<TD width="6%"><a href="?action=Outbox"><img src="Images/M_outbox.gif" width="40" height="40" border="0"></a></TD>
									<TD width="6%"><a href="?action=Recycle"><img src="Images/M_recycle.gif" width="40" height="40" border="0"></a></TD>
									<TD width="6%"><a href="User_AddressList.asp"><img src="Images/M_address.gif" width="40" height="40" border="0"></a></TD>
									<TD width="2%"><span class="f41"><a href="User_WriteMessage.asp"><img src="Images/m_write.gif" width="40" height="40" border="0"></a></span></TD>
									<%
									Dim SumRsObj,TotleSQL,UnTotle,UnTotle1
									Set SumRsObj = Server.CreateObject(G_FS_RS)
									TotleSQL = "Select sum(LenContent) from FS_Message where MeRead='"& RsUserObj("MemName") &"' and IsDelR = 0"
									SumRsObj.Open TotleSQL,Conn,1,3
									If Not SumRsObj.Eof then
										UnTotle=SumRsObj(0)/(1024*50)*100
									Else
										UnTotle=0
									End if
									If IsNull(UnTotle) then UnTotle=0
									%>
									<TD width="68%"><div align="center">空间已经使用： 
										<% = CInt(UnTotle)%>
										%</div></TD>
								  </TR>
								</TBODY>
							  </TABLE>
							  <hr size="1" noshade>
							<%
							If Request("action")="Inbox" Then
								Call Inbox()
							ElseIf Request("action")="Outbox"  Then
								Call OutBox()
							ElseIf Request("action")="Recycle" Then
								Call Recycle()
							Else
								Call Inbox()
							End if  
							%>
							<%
							Sub Inbox()
							%>
							  <strong>收件箱</strong><br>
							  <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
								<form name="form1" method="post" action="">
								  <tr bgcolor="#EFEFEF"> 
									<td width="38"> <div align="center"><strong>已读</strong></div></td>
									<td width="99"> <div align="center"><strong>发件人</strong></div></td>
									<td width="387"><div align="center"><strong>主题</strong></div></td>
									<td width="183"><div align="center"><strong>日期</strong></div></td>
									<td width="117"><div align="center"><strong>大小</strong></div></td>
									<td width="43"><div align="center"><strong>操作</strong></div></td>
								  </tr>
								  <%
									Dim strpage,RsInboxObj,InboxSQL,Select_count,Select_pagecount
									strpage=request("page")
									if len(strpage)=0 then
										strpage="1"
									end if
									Set RsInboxObj = Server.CreateObject(G_FS_RS)
									InboxSQL="Select * from FS_Message where MeRead='"& RsUserObj("MemName") &"' and IsDelR=0 and IsRecyle=0 order by MeID desc"
									RsInboxObj.Open InboxSQL,Conn,1,1
									If Not RsInboxObj.eof then
										RsInboxObj.pagesize=20
										RsInboxObj.absolutepage=cint(strpage)
										Select_count=RsInboxObj.recordcount
										Select_pagecount=RsInboxObj.pagecount
										For i=1 to RsInboxObj.pagesize
											If RsInboxObj.eof then
												exit for
											End if
												If RsInboxObj.eof then
													exit for
												End if
									%>
								  <tr bgcolor="#FFFFFF"> 
									<td> <%
										Dim Strs,Strs1
										If RsInboxObj("ReadTF")=0 then
											Strs="<b>"
											Strs1="</b>"
										%> <img src="Images/Read.gif" alt="未读" width="21" height="14" style="CURSOR: hand"> 
																	  <%
										Else
											Strs=""
											Strs1=""
										%> <img src="Images/Readed.gif" alt="已读" width="21" height="14" style="CURSOR: hand">	
										<%
										End If
										%>
									</td>
									<td><% = Strs %> <a href=ReadUser.Asp?UserName=<% = RsInboxObj("MeFrom") %> target="_blank"> 
									  <% = RsInboxObj("MeFrom")%></a> <% = Strs1 %></td>
									<td><% = Strs %> <a href=User_ReadMessage.Asp?id=<% = RsInboxObj("MeId") %>> 
									  <% = left(RsInboxObj("MeTiTle"),20)%>
									  </a> <% = Strs1 %></td>
									<td><% = Strs %> <% = RsInboxObj("FromDate")%> <% = Strs1 %></td>
									<td><% = Strs %> <% = RsInboxObj("LenContent")%>
									  Byte <% = Strs1 %></td>
									<td><input name="InboxID" type="checkbox" id="InboxID" value="<% = RsInboxObj("MeID")%>"></td>
								  </tr>
									<%
										RsInboxObj.MoveNext
									Next
									%>
								  <tr bgcolor="#FFFFFF"> 
									<td colspan="6"><div align="right">
										<input type="submit" name="Submit" value="删除选中短信到回收站">
										<input name="action" type="hidden" id="action" value="delInbox">
									  </div></td>
								  </tr>
	
								</form>
							  </table>
	<%
								  Response.write"<br>&nbsp;共<b>"& select_pagecount &"</b>页<b>" & select_count &"</b>条记录，本页是第<b>"& strpage &"</b>页。"
										If int(strpage)>1 Then
											Response.Write"&nbsp;<a href=?page=1>第一页</a>&nbsp;"
											Response.Write"&nbsp;<a href=?page="&cstr(cint(strpage)-1)&">上一页</a>&nbsp;"
										End If
											If int(strpage)<select_pagecount Then
											Response.Write"&nbsp;<a href=?page="&cstr(cint(strpage)+1)&">下一页</a>"
											Response.Write"&nbsp;<a href=?page="& select_pagecount &">最后一页</a>&nbsp;"
										End If
								  Response.Write("<br>")
								Else
									response.Write("<tr><td colspan=""6"" bgcolor=""#FFFFFF"">没有短信</td></tr></table>")
								End if
								Set RsInboxObj=nothing
	End Sub
	%>                      
	<%
	Sub OutBox()
	%>
							  <strong>发件箱</strong> 
							  <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
								<form name="form1" method="post" action="">
								  <tr bgcolor="#EFEFEF"> 
									<td width="38"> <div align="center"><strong>已读</strong></div></td>
									<td width="99"> <div align="center"><strong>发件人</strong></div></td>
									<td width="387"><div align="center"><strong>主题</strong></div></td>
									<td width="183"><div align="center"><strong>日期</strong></div></td>
									<td width="117"><div align="center"><strong>大小</strong></div></td>
									<td width="43"><div align="center"><strong>操作</strong></div></td>
								  </tr>
								  <%
									Dim strpage,RsInboxObj,InboxSQL,Select_count,Select_pagecount
									strpage=request("page")
									if len(strpage)=0 then
										strpage="1"
									end if
									Set RsInboxObj = Server.CreateObject(G_FS_RS)
									InboxSQL="Select * from FS_Message where MeFrom='"& RsUserObj("MemName") &"' and isSend=1 and IsRecyle=0 order by MeID desc"
									RsInboxObj.Open InboxSQL,Conn,1,1
									If Not RsInboxObj.eof then
										RsInboxObj.pagesize=20
										RsInboxObj.absolutepage=cint(strpage)
										Select_count=RsInboxObj.recordcount
										Select_pagecount=RsInboxObj.pagecount
										For i=1 to RsInboxObj.pagesize
											If RsInboxObj.eof then
												exit for
											End if
												If RsInboxObj.eof then
													exit for
												End if
									%>
								  <tr bgcolor="#FFFFFF"> 
									<td> 
									  <%
										Dim Strs,Strs1
										If RsInboxObj("ReadTF")=0 then
											Strs="<b>"
											Strs1="</b>"
										%>
									  <img src="Images/Read.gif" alt="未读" width="21" height="14" style="CURSOR: hand"> 
									  <%
										Else
											Strs=""
											Strs1=""
										%>
									  <img src="Images/Readed.gif" alt="已读" width="21" height="14" style="CURSOR: hand">	
									  <%
										End If
										%>
									</td>
									<td> 
									  <% = Strs %>
									  <a href=ReadUser.Asp?UserName=<% = RsInboxObj("MeFrom") %> target="_blank"> 
									  <% = RsInboxObj("MeFrom")%>
									  </a> 
									  <% = Strs1 %>
									</td>
									<td> 
									  <% = Strs %>
									  <a href=User_ReadMessage_Re.Asp?id=<% = RsInboxObj("MeId") %>> 
									  <% = left(RsInboxObj("MeTiTle"),20)%>
									  </a> 
									  <% = Strs1 %>
									</td>
									<td> 
									  <% = Strs %>
									  <% = RsInboxObj("FromDate")%>
									  <% = Strs1 %>
									</td>
									<td> 
									  <% = Strs %>
									  <% = RsInboxObj("LenContent")%>
									  Byte 
									  <% = Strs1 %>
									</td>
									<td><input name="OutBoxID" type="checkbox" id="OutBoxID" value="<% = RsInboxObj("MeID")%>"></td>
								  </tr>
								  <%
										RsInboxObj.MoveNext
									Next
									%>
								  <tr bgcolor="#FFFFFF"> 
									<td colspan="6"><div align="right"> 
										<input type="submit" name="Submit22" value="彻底删除">
										<input name="action" type="hidden" id="action3" value="OutBox">
									  </div></td>
								  </tr>
								</form>
							  </table>
							  <%
								  Response.write"<br>&nbsp;共<b>"& select_pagecount &"</b>页<b>" & select_count &"</b>条记录，本页是第<b>"& strpage &"</b>页。"
										If int(strpage)>1 Then
											Response.Write"&nbsp;<a href=?page=1&action=OutBox>第一页</a>&nbsp;"
											Response.Write"&nbsp;<a href=?page="&cstr(cint(strpage)-1)&"&action=OutBox>上一页</a>&nbsp;"
										End If
											If int(strpage)<select_pagecount Then
											Response.Write"&nbsp;<a href=?page="&cstr(cint(strpage)+1)&"&action=OutBox>下一页</a>"
											Response.Write"&nbsp;<a href=?page="& select_pagecount &"&action=OutBox>最后一页</a>&nbsp;"
										End If
								  Response.Write("<br>")
								Else
									response.Write("<tr><td colspan=""6"" bgcolor=""#FFFFFF"">没有短信</td></tr></table>")
								End if
								Set RsInboxObj=nothing
								End Sub
								%>
							  <%
	Sub Recycle()
	%>
							  <strong>废件箱</strong> 
							  <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
								<form name="form1" method="post" action="">
								  <tr bgcolor="#EFEFEF"> 
									<td width="38"> <div align="center"><strong>已读</strong></div></td>
									<td width="99"> <div align="center"><strong>发件人</strong></div></td>
									<td width="387"><div align="center"><strong>主题</strong></div></td>
									<td width="183"><div align="center"><strong>日期</strong></div></td>
									<td width="117"><div align="center"><strong>大小</strong></div></td>
									<td width="43"><div align="center"><strong>操作</strong></div></td>
								  </tr>
								  <%
									Dim strpage,RsInboxObj,InboxSQL,Select_count,Select_pagecount,Strs,Strs1
									strpage=request("page")
									if len(strpage)=0 then
										strpage="1"
									end if
									Set RsInboxObj = Server.CreateObject(G_FS_RS)
									InboxSQL="Select * from FS_Message where MeRead='"& RsUserObj("MemName") &"' and IsDelR=0 and IsRecyle=1 order by MeID desc"
									RsInboxObj.Open InboxSQL,Conn,1,1
									If Not RsInboxObj.eof then
										RsInboxObj.pagesize=20
										RsInboxObj.absolutepage=cint(strpage)
										Select_count=RsInboxObj.recordcount
										Select_pagecount=RsInboxObj.pagecount
										For i=1 to RsInboxObj.pagesize
											If RsInboxObj.eof then
												exit for
											End if
												If RsInboxObj.eof then
													exit for
												End if
									%>
								  <tr bgcolor="#FFFFFF"> 
									<td> 
									  <%
										If RsInboxObj("ReadTF")=0 then
											Strs="<b>"
											Strs1="</b>"
										%>
									  <img src="Images/Read.gif" alt="未读" width="21" height="14" style="CURSOR: hand"> 
									  <%
										Else
											Strs=""
											Strs1=""
										%>
									  <img src="Images/Readed.gif" alt="已读" width="21" height="14" style="CURSOR: hand">	
									  <%
										End If
										%>
									</td>
									<td>
									  <% = Strs %>
									  <a href=ReadUser.Asp?UserName=<% = RsInboxObj("MeFrom") %> target="_blank"> 
									  <% = RsInboxObj("MeFrom")%>
									  </a> 
									  <% = Strs1 %>
									</td>
									<td>
									  <% = Strs %>
									  <a href=User_ReadMessage.Asp?id=<% = RsInboxObj("MeId") %>> 
									  <% = left(RsInboxObj("MeTiTle"),20)%>
									  </a> 
									  <% = Strs1 %>
									</td>
									<td>
									  <% = Strs %>
									  <% = RsInboxObj("FromDate")%>
									  <% = Strs1 %>
									</td>
									<td>
									  <% = Strs %>
									  <% = RsInboxObj("LenContent")%>
									  Byte 
									  <% = Strs1 %>
									</td>
									<td><input name="RecycleboxID" type="checkbox" id="RecycleboxID" value="<% = RsInboxObj("MeID")%>"></td>
								  </tr>
								  <%
										RsInboxObj.MoveNext
									Next
									%>
								  <tr bgcolor="#FFFFFF"> 
									<td colspan="6"><div align="right"> 
										<input type="submit" name="Submit2" value="彻底删除">
										<input name="action" type="hidden" id="action" value="RecycleInbox">
									  </div></td>
								  </tr>
								</form>
							  </table>
	<%
								  Response.write"<br>&nbsp;共<b>"& select_pagecount &"</b>页<b>" & select_count &"</b>条记录，本页是第<b>"& strpage &"</b>页。"
										If int(strpage)>1 Then
											Response.Write"&nbsp;<a href=?page=1&action=Recycle>第一页</a>&nbsp;"
											Response.Write"&nbsp;<a href=?page="&cstr(cint(strpage)-1)&"&action=Recycle>上一页</a>&nbsp;"
										End If
											If int(strpage)<select_pagecount Then
											Response.Write"&nbsp;<a href=?page="&cstr(cint(strpage)+1)&"&action=Recycle>下一页</a>"
											Response.Write"&nbsp;<a href=?page="& select_pagecount &"&action=Recycle>最后一页</a>&nbsp;"
										End If
								  Response.Write("<br>")
								Else
									response.Write("<tr><td colspan=""6"" bgcolor=""#FFFFFF"">没有短信</td></tr></table>")
								End if
								Set RsInboxObj=nothing
								%>
	<%
	End Sub
	%>                      
						  
							  <table width="100%" height="26" border="0" cellpadding="5" cellspacing="0">
								<tr>
									
								  <td height="26">&nbsp;</td>
								  </tr>
								</table></TD>
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
End if
RsConfigObj.Close
Set RsConfigObj = Nothing
RsUserObj.close
Set RsUserObj=nothing
Set Conn=nothing
%>

