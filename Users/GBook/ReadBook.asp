<% Option Explicit %>
<!--#include file="../../Inc/Function.asp" -->
<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Const.asp" -->
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
	Set RsConfigObj = Conn.Execute("Select SiteName,UserConfer,Copyright,isEmail,isChange,IsShop,QPoint,MaxContent from FS_Config")
	Set DBC = Nothing
%>
<!--#include file="../Comm/User_Purview.Asp" -->
<%
If Request.Form("action")="add" then
		If trim(request.form("Content"))="" then
			Response.Write("<script>alert(""请填写回复内容"&CopyRight&""");location=""javascript:history.back(1)"";</script>")  
			Response.End
		End if
		If Len(request.form("Content"))>RsConfigObj("MaxContent") then
			Response.Write("<script>alert(""内容不能超过"& RsConfigObj("MaxContent") &"字符"&CopyRight&""");location=""javascript:history.back(1)"";</script>")  
			Response.End
		End if
		If  Cint(Session("MemID"))=0 then
			Response.Write("<script>alert(""错误的权限！！！"&CopyRight&""");location=""javascript:history.back(1)"";</script>")  
			Response.End
		End if
	  Dim Rs,Sql1
	  Set Rs = server.createobject(G_FS_RS)
	  Sql1 = "select * from FS_GBook where 1=0"
	  Rs.open sql1,conn,1,3
	  Rs.addnew
	  Rs("Content")=Trim(NoCSSHackContent(Request.Form("Content")))
	  Rs("AddTime")=Now()
	  Rs("UserID")=Session("MemID")
	  Rs("FaceNum")=NoCSSHackInput(Replace(request.form("FaceNum"),"'",""))
	  Rs("isQ")=0
	  Rs("isLock")=0
	  Rs("isAdmin")=0
	  Rs("Orders")=2
	  Rs("EditQ")=""
	  Rs("QID")=NoCSSHackInput(Replace(Request.form("QID"),"'",""))
	  Rs.update
	  '更新恢复帖子
	   Conn.execute("Update FS_GBook Set isQ = 1,Qtime="&StrSqlDate&" where id="&Replace(Replace(Request.form("QID"),"'",""),Chr(39),""))
	  '增加积分
	   Conn.execute("Update FS_Members Set Point = Point+"&RsConfigObj("QPoint")&" where id="&Replace(Replace(Session("MemId"),"'",""),Chr(39),""))
	  Response.Write("<script>alert(""回复成功"&CopyRight&""");location=""ReadBook.asp?ID="& Replace(request.form("QID"),"'","") &""";</script>") 
	  Response.End
	  Rs.close
	  Set rs=nothing
End if
iF Request("Action")="Del" then
	Dim GBListObj
	Set GBListObj = Conn.execute("Select ID,UserID From FS_GBook where ID="&Replace(Replace(Request("Id"),"'",""),Chr(39),""))
	If Cint(GBListObj("UserID"))<>Cint(Session("MemID")) Then
		Response.Write("<script>alert(""您没权限删除此帖"&CopyRight&""");location=""javascript:history.back(1)"";</script>")  
		Response.End
	End if
	Conn.execute("Delete From FS_GBook where id="&Replace(Replace(Request("Id"),"'",""),Chr(39),""))
	'扣除会员积分
	If Request("GetAction")="1" then
		Response.Write("<script>alert(""删除成功"&CopyRight&""");location=""All_GBook.asp"";</script>") 
	Else
		Conn.execute("Update FS_Members Set Point = Point-"&RsConfigObj("QPoint")&" where id="&Replace(Replace(Session("MemId"),"'",""),Chr(39),""))
		Response.Write("<script>alert(""删除成功"&CopyRight&""");location=""ReadBook.asp?id="&Request("sid")&""";</script>") 
	End if 
	Response.End
End if
Dim NewsContent
NewsContent = Replace(Replace(Request.Form("Content"),"""","%22"),"'","%27")
Dim RsModifyObj,ModifySQL
  Set RsModifyObj = server.createobject(G_FS_RS)
  ModifySQL = "select * from FS_GBook where ID="&Replace(Replace(Request("Id"),"'",""),Chr(39),"")
  RsModifyObj.open ModifySQL,conn,1,1
  If RsModifyObj.eof then
	  Response.Write("<script>alert(""找不到记录！！"&CopyRight&""");location=""javascript:history.back()"";</script>") 
	  Response.End
  End if
  iF Cint(RsModifyObj("IsAdmin"))=1 then
  	If RsModifyObj("UserID")<>Session("MemID") Then
	  Response.Write("<script>alert(""此帖只有管理员才能查看！！"&CopyRight&""");location=""javascript:history.back()"";</script>") 
	  Response.End
	 End if
  End if
Dim MemberObj
Set MemberObj = Conn.execute("Select Sex,HeadPic,MemName,OICQ,Email From FS_Members Where id="&Replace(Replace(RsModifyObj("Userid"),"'",""),Chr(39),""))
%>
<HTML><HEAD>
<TITLE><%=RsConfigObj("SiteName")%> >> 会员中心</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="../Css/UserCSS.css" type=text/css  rel=stylesheet>
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
                              src="../images/Favorite.OnArrow.gif" border=0></TD>
                      <TD 
class=f4>发表帖子</TD>
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
                    <table width="75%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td height="3"></td>
                      </tr>
                    </table>
                    
                  <table width="100%" border="0" cellspacing="0" cellpadding="5">
                    <tr>
                      
                    <td width="62%"><a href="GBook.asp">我发表的帖子</a> ｜ <a href="All_GBook.asp">帖子查看</a> 
                      ｜ <a href="Write_GBook.asp"><font color="#FF0000">发表帖子</font></a> 
                      ｜ <a href="GBook.asp?Action=Q">已回复的帖子</a> ｜ <a href="GBook.asp?Action=Q"></a><a href="GBook.asp?Action=UnQ">未回复的帖子</a></td>
                  <form name="form2" method="post" action="all_GBook.asp">
                      <td width="38%"><input name="Keyword" type="text" id="Keyword">
                        <input type="submit" name="Submit2" value="搜索"> </td>
                    </form>
                    </tr>
                  </table>
				  <%iF RsModifyObj("IsLock")=0 then%>
                <table width="96%" border="0" cellpadding="0" cellspacing="0">
                  <tr> 
                    <td><a href="Write_GBook.asp"><img src="Images/postnew.gif" width="85" height="26" border="0"></a>　<a href="ReadBook.asp?Id=<% = RsModifyObj("id")%>&QAction=Q#QU"><img src="Images/mreply.gif" width="85" height="26" border="0"></a></td>
                  </tr>
                </table>
				<%End if%>
                <table width="100%" height="114" border="0" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                  <tr bgcolor="#FFFFFF"> 
                    <td width="15%" valign="top" bgcolor="#EFEFEF">
					<%
					If RsModifyObj("UserID")<>0  then
					%>
					<table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr> 
                          <td width="71%"> <strong> 
                            <a href=../ReadUser.asp?UserName=<% = MemberObj("MemName")%>><% = MemberObj("MemName")%></a>
                            </strong></td>
                          <td width="29%"> <%
						  If MemberObj("Sex") =0 then
						  %>
                            <img src="Images/Male.gif" alt="帅哥哦" width="23" height="21"> 
                            <%Else%>
                            <img src="Images/FeMale.gif" alt="美女哦" width="23" height="21"> 
                            <%End if%></td>
                        </tr>
                      </table>
                      <div align="center"> 
                        <hr size="1" noshade color="#CCCCCC">
                        <%If Len(MemberObj("HeadPic"))>5 then%>
                        <img src=../<% = MemberObj("HeadPic")%>> 
                        <%Else%>
                        <table width="0" border="0" cellpadding="0" cellspacing="0" bgcolor="#F0F0F0">
                          <tr> 
                            <td bgcolor="#FFFFFF"><img src="../images/noHeadPic.jpg" width="50" height="50" border="0"></td>
                          </tr>
                        </table>
                        <%End if%>
					  <%Else%>
                        <strong><font color="#990033">管理员</font></strong> 
                        <%End if%>                      </div>
                      <br> <% = RsModifyObj("addtime")%>
</td>
                    <td width="85%" valign="top"> <table width="100%" border="0" cellspacing="0" cellpadding="3">
                        <tr> 
                          <td colspan="2" valign="top">
						  <%If RsModifyObj("UserID")<>0  then%>
						  <table width="60%" border="0" cellspacing="0" cellpadding="3">
                              <tr> 
                                <td width="86"> 
                                  <div align="center">
                                    <%
						if Trim(MemberObj("OICQ"))<>"" then
							Dim sOICQ
						    sOICQ ="<a target=blank href=http://wpa.qq.com/msgrd?V=1&Uin="& MemberObj("OICQ") &"&Site="& RsConfigObj("SiteName") &"&Menu=yes><img border=""0"" SRC=http://wpa.qq.com/pa?p=1:"& MemberObj("OICQ") &":8 alt=""点击这里给"& MemberObj("OICQ") &"发消息""></a>"
							Response.Write sOICQ
						Else
							Response.Write("没有QQ")
						End if
						%>
                                  </div></td>
                                <td width="48"><a href="../User_AddressList.asp?UserName=<% = MemberObj("MemName")%>"><img src="Images/friend.gif" alt="加为好友" width="48" height="18" border="0"></a></td>
                                <td width="45"><a href="../ReadUser.asp?UserName=<% = MemberObj("MemName")%>"><img src="Images/profile.gif" alt="查看信息" width="45" height="18" border="0"></a></td>
                                <td width="45"><a href="mailto:<%=MemberObj("Email")%>"><img src="Images/email.gif" width="45" height="18" border="0"></a></td>
                                <td width="113"><a href="../User_WriteMessage.asp?UserName=<% = MemberObj("MemName")%>"><img src="Images/message.gif" width="48" height="18" border="0"></a></td>
                              </tr>
                              <tr bgcolor="#D0D0D0"> 
                                <td height="1" colspan="5"></td>
                              </tr>
                            </table><%End if%></td>
                        </tr>
                        <tr> 
                          <td width="4%" valign="top"><img src="Images/face<% = RsModifyObj("FaceNum")%>.gif" width="22" height="22"></td>
                          <td width="96%" valign="bottom"><strong> 
                            <% = RsModifyObj("Title")%>
                            </strong></td>
                        </tr>
                        <tr> 
                          <td height="29">&nbsp;</td>
                          <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                              <tr> 
                                <td height="5"></td>
                              </tr>
                            </table>
                            <% = RsModifyObj("Content")&RsModifyObj("EditQ")%> </td>
                        </tr>
                        <tr> 
                          <td>&nbsp;</td>
                          <td><div align="right"> 
                              <%If RsModifyObj("UserID")=Session("MemID") Then %>
                              <a href="Modify_GBook.asp?ID=<%=RsModifyObj("id")%>">[编辑此帖子]</a> 
                              <a href="ReadBook.asp?ID=<%=RsModifyObj("id")%>&Action=Del&GetAction=1"  onClick="return Cim()">[删除此帖子]</a> 
                              <%Else%>
                              <font color="#999999">[编辑此帖子]</font> 
                              <%End if%>
                            </div></td>
                        </tr>
                      </table></td>
                  </tr>
					<%
					Dim RsQModifyObj,QModifySQL
					Dim RsCon,strpage,select_count,select_pagecount
					strpage=request.querystring("page")
					if len(strpage)=0 then
						strpage="1"
					end if
					Set RsQModifyObj = server.createobject(G_FS_RS)
					QModifySQL = "select * from FS_GBook where QID="&Replace(Replace(RsModifyObj("Id"),"'",""),Chr(39),"")
					RsQModifyObj.open QModifySQL,conn,1,1
					If Not RsQModifyObj.eof then
							RsQModifyObj.pagesize=10
							RsQModifyObj.absolutepage=cint(strpage)
							select_count=RsQModifyObj.recordcount
							select_pagecount=RsQModifyObj.pagecount
							for i=1 to RsQModifyObj.pagesize
							if RsQModifyObj.eof then
								exit for
							end if
					if i mod 2 <> 0 then
					%>
                    <tr bgcolor="#EEEEEE"> 
                      <%Else%>
                    <tr bgcolor="#FFFFFF"> 
                      <%End If%>
                    <td valign="top" bgcolor="#EFEFEF"> 
					<%
					If RsQModifyObj("UserID")<>0  then
					%>
                      <table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr> 
                          <td width="71%"> <strong> 
                            <%
					  	Dim QMemberObj
						Set QMemberObj = Conn.execute("Select MemName,Sex,OICQ,Email,HeadPic From FS_Members Where id="&Replace(Replace(RsQModifyObj("UserID"),"'",""),Chr(39),""))
						If Not QMemberObj.eof then
							Response.Write("<a href=../ReadUser.Asp?UserName="&QMemberObj("MemName")&">"& QMemberObj("MemName")&"</a>")
						Else
							Response.Write("用户已被删除")
						End if
					  %>
                            </strong></td>
                          <td width="29%"> 
                            <%
						  If QMemberObj("Sex") =0 then
						  %>
                            <img src="Images/Male.gif" alt="帅哥哦" width="23" height="21"> 
                            <%Else%>
                            <img src="Images/FeMale.gif" alt="美女哦" width="23" height="21"> 
                            <%End if%>
                          </td>
                        </tr>
                      </table>
                      <div align="center"> 
                        <hr size="1" noshade color="#CCCCCC">
                        <%If Len(QMemberObj("HeadPic"))>5 then%>
                        <img src=../<% = QMemberObj("HeadPic")%>> 
                        <%Else%>
                        <table width="0" border="0" cellpadding="0" cellspacing="0" bgcolor="#F0F0F0">
                          <tr> 
                            <td bgcolor="#FFFFFF"><img src="../images/noHeadPic.jpg" width="50" height="50" border="0"></td>
                          </tr>
                        </table>
                      </div>
                      <div align="center"></div>
                      <div align="center">
                        <%End if
						Else%>
                        <strong><font color="#990000">管理员</font></strong> 
                        <%End if%>
                        <br>
                        <%=RsQModifyObj("Addtime")%><br>
                      </div></td>
                    <td valign="top"> 
                      <%If RsQModifyObj("UserID")<>0  then%>
                      <table width="60%" border="0" cellspacing="0" cellpadding="3">
                        <tr> 
                          <td width="30%"> <div align="center"> 
                              <%
						if Trim(QMemberObj("OICQ"))<>"" then
							Dim sOICQ1
						    sOICQ1 ="<a target=blank href=http://wpa.qq.com/msgrd?V=1&Uin="& QMemberObj("OICQ") &"&Site="& RsConfigObj("SiteName") &"&Menu=yes><img border=""0"" SRC=http://wpa.qq.com/pa?p=1:"& QMemberObj("OICQ") &":8 alt=""点击这里给"& QMemberObj("OICQ") &"发消息""></a>"
							Response.Write sOICQ1
						Else
							Response.Write("没有QQ")
						End if
						%>
                            </div></td>
                          <td width="14%"><a href="../User_AddressList.asp?UserName=<% = QMemberObj("MemName")%>"><img src="Images/friend.gif" alt="加为好友" width="48" height="18" border="0"></a></td>
                          <td width="13%"><a href="../ReadUser.asp?UserName=<% = QMemberObj("MemName")%>"><img src="Images/profile.gif" alt="查看信息" width="45" height="18" border="0"></a></td>
                          <td width="12%"><a href="mailto:<%=QMemberObj("Email")%>"><img src="Images/email.gif" width="45" height="18" border="0"></a></td>
                          <td width="31%"><a href="../User_WriteMessage.asp?UserName=<% = QMemberObj("MemName")%>"><img src="Images/message.gif" width="48" height="18" border="0"></a></td>
                        </tr>
                        <tr bgcolor="#D0D0D0"> 
                          <td height="1" colspan="5"></td>
                        </tr>
                      </table>
                      <%End if%>
                      <table width="100%" height="107" border="0" cellpadding="0" cellspacing="0">
                        <tr> 
                          <td width="4%" height="30" valign="top"><img src="Images/face<% = RsQModifyObj("FaceNum")%>.gif" width="22" height="22"></td>
                          <td width="96%" valign="top"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                <td height="8"></td>
                              </tr>
                            </table><% = RsQModifyObj("Content")&RsQModifyObj("EditQ")%>
                          </td>
                        </tr>
                        <tr> 
                          <td>&nbsp;</td>
                          <td valign="bottom">
<div align="right"> 
                              <%If RsQModifyObj("UserID")=Session("MemID") Then %>
                              <a href="Modify_GBook.asp?ID=<%=RsQModifyObj("id")%>&GetAction=oper&Sid=<%=Request("id")%>">[编辑此帖子]</a> 
                              <a href="ReadBook.asp?ID=<%=RsQModifyObj("id")%>&Action=Del&Sid=<%=Request("id")%>&GetAction=2"  onClick="return Cim()">[删除此帖子]</a> 
                              <%Else%>
                              <font color="#999999">[编辑此帖子]</font> 
                              <%End if%>
                            </div></td>
                        </tr>
                      </table></td>
                  </tr>
					<%
					     RsQModifyObj.MoveNext
					 Next
					%>
                  <tr bgcolor="#FFFFFF">
                    <td valign="top" bgcolor="#EFEFEF">&nbsp;</td>
                    <td valign="top">
					<%
						Response.write"<br>&nbsp;共<b>"& select_pagecount &"</b>页<b>" & select_count &"</b>个回复帖子，本页是第<b>"& strpage &"</b>页。"
						if int(strpage)>1 then
							Response.Write"&nbsp;<a href=?page=1&ID="&RsModifyObj("id")&">第一页</a>&nbsp;"
							Response.Write"&nbsp;<a href=?page="&cstr(cint(strpage)-1)&"&ID="&RsModifyObj("id")&">上一页</a>&nbsp;"
						end if
						if int(strpage)<select_pagecount then
							Response.Write"&nbsp;<a href=?page="&cstr(cint(strpage)+1)&"&ID="&RsModifyObj("id")&">下一页</a>"
							Response.Write"&nbsp;<a href=?page="& select_pagecount &"&ID="&RsModifyObj("id")&">最后一页</a>&nbsp;"
						end if
						Response.Write"<br>"
						RsQModifyObj.close
						Set RsQModifyObj=nothing
					%> 
					</td>
                  </tr>
				  <%End if%>
                </table></TD>
            </TR>
          </TBODY>
        </TABLE>
		<%iF RsModifyObj("IsLock")=0 then%>
        <table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td><a href="Write_GBook.asp"><img src="Images/postnew.gif" width="85" height="26" border="0"></a>　<a href="ReadBook.asp?Id=<% = RsModifyObj("id")%>&QAction=Q#QU"><img src="Images/mreply.gif" width="85" height="26" border="0"></a></td>
          </tr>
        </table>
        <%
		End if
If Request("QAction")="Q" Then
	Call QuickQ()
End if
Sub QuickQ()
%>
        <a name="QU"></a> 
        <table width="97%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
          <form action="ReadBook.asp" method="POST" name="NewsForm">
            <tr bgcolor="#FFFFFF"> 
              <td height="19" colspan="2" bgcolor="#EFEFEF"> 
                <div align="left"><strong>快速回复帖子</strong></div></td>
            </tr>
            <tr bgcolor="#FFFFFF"> 
              <td width="16%" bgcolor="#F3F3F3"> 
                <div align="right">用户名：</div></td>
              <td width="84%"> <input name="textfield" type="text" value="<%=Session("MemName")%>" readonly></td>
            </tr>
            <tr bgcolor="#FFFFFF"> 
              <td bgcolor="#F3F3F3"> 
                <div align="right">表情：</div></td>
              <td> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr> 
                    <td> <input name="FaceNum" type="radio" value="1" checked> 
                      <img src="Images/face1.gif" width="22" height="22"> </td>
                    <td> <input type="radio" name="FaceNum" value="2"> <img src="Images/face2.gif" width="22" height="22"> 
                    </td>
                    <td> <input type="radio" name="FaceNum" value="3"> <img src="Images/face3.gif" width="22" height="22"> 
                    </td>
                    <td> <input type="radio" name="FaceNum" value="4"> <img src="Images/face4.gif" width="22" height="22"> 
                    </td>
                    <td> <input type="radio" name="FaceNum" value="5"> <img src="Images/face5.gif" width="22" height="22"> 
                    </td>
                    <td> <input type="radio" name="FaceNum" value="6"> <img src="Images/face6.gif" width="22" height="22"></td>
                    <td> <input type="radio" name="FaceNum" value="7"> <img src="Images/face7.gif" width="22" height="22"> 
                    </td>
                    <td> <input type="radio" name="FaceNum" value="8"> <img src="Images/face8.gif" width="22" height="22"> 
                    </td>
                    <td> <input type="radio" name="FaceNum" value="9"> <img src="Images/face9.gif" width="22" height="22"></td>
                  </tr>
                  <tr> 
                    <td> <input type="radio" name="FaceNum" value="10"> <img src="Images/face10.gif" width="22" height="22"></td>
                    <td> <input type="radio" name="FaceNum" value="11"> <img src="Images/face11.gif" width="22" height="22"> 
                    </td>
                    <td> <input type="radio" name="FaceNum" value="12"> <img src="Images/face12.gif" width="22" height="22"></td>
                    <td> <input type="radio" name="FaceNum" value="13"> <img src="Images/face13.gif" width="22" height="22"> 
                    </td>
                    <td> <input type="radio" name="FaceNum" value="14"> <img src="Images/face14.gif" width="22" height="22"></td>
                    <td> <input type="radio" name="FaceNum" value="15"> <img src="Images/face15.gif" width="22" height="22"></td>
                    <td> <input type="radio" name="FaceNum" value="16"> <img src="Images/face16.gif" width="22" height="22"></td>
                    <td> <input type="radio" name="FaceNum" value="17"> <img src="Images/face17.gif" width="22" height="22"> 
                    </td>
                    <td> <input type="radio" name="FaceNum" value="18"> <img src="Images/face18.gif" width="22" height="22"> 
                    </td>
                  </tr>
                </table></td>
            </tr>
            <tr bgcolor="#FFFFFF"> 
              <td colspan="2"> <div align="right"></div>
                <iframe id='NewsContent' src='../Editer/BookQNewsEditer.asp' frameborder=0 scrolling=no width='100%' height='200'></iframe></td>
            </tr>
            <tr bgcolor="#FFFFFF"> 
              <td>&nbsp;</td>
              <td> <input name="submitggg" type="button" onClick="SubmitFun();" value="回复帖子"> 
                <input name="reset" type="reset" value="重新填写"> <input name="Content" type="hidden" id="Content" value="<% = NewsContent %>">
                <input name="QID" type="hidden" id="QID" value="<% = Request("ID") %>">
                <input name="action" type="hidden" id="action" value="add"> </td>
            </tr>
          </form>
        </table>
<%End Sub%>
        </TD>
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
Set RsModifyObj =Nothing
Set MemberObj =Nothing
RsConfigObj.Close
Set RsConfigObj = Nothing
Set Conn=nothing
%>
<script>
function SubmitFun()
{
	frames["NewsContent"].SaveCurrPage();
	var TempContentArray=frames["NewsContent"].NewsContentArray;
	document.NewsForm.Content.value='';
	for (var i=0;i<TempContentArray.length;i++)
	{
		if (TempContentArray[i]!='')
		{
			if (document.NewsForm.Content.value=='') document.NewsForm.Content.value=TempContentArray[i];
			else document.NewsForm.Content.value=document.NewsForm.Content.value+'[Page]'+TempContentArray[i];
		} 
	}
	document.NewsForm.submit();
}
function Cim(){
	if (window.confirm('您确定要删除吗?删除不可逆！！')){
	 	return true;
	 } 
	 return false;		
}
</script>