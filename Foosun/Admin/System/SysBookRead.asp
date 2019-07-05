<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<%
'==============================================================================
'软件名称：FoosunShop System Form FoosunCMS
'当前版本：Foosun Content Manager System 3.0 系列
'最新更新：2004.12
'==============================================================================
'商业注册联系：028-85098980-601,602 技术支持：028-85098980-605、607,客户支持：608
'产品咨询QQ：159410,394226379,125114015,655071
'技术支持:所有程序使用问题，请提问到bbs.foosun.net我们将及时回答您
'程序开发：风讯开发组 & 风讯插件开发组
'论坛支持：风讯在线论坛(http://bbs.foosun.net)
'官方网站：www.Foosun.net  演示站点：test.cooin.com    
'网站建设专区：www.cooin.com
'==============================================================================
'免费版本请在新闻首页保留版权信息，并做上本站LOGO友情连接
'==============================================================================
Dim DBC,Conn,sRootDir
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
Dim RsAdminConfigObj,IsShowSave
Set RsAdminConfigObj = Conn.Execute("Select SiteName,UserConfer,Copyright,isEmail,isChange,IsShop,MaxContent,QPoint from FS_Config")
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P070700") then Call ReturnError1()
if Not JudgePopedomTF(Session("Name"),"P070702") then Call ReturnError1()

If Request.Form("action")="add" then
		If trim(request.form("Content"))="" then
			Response.Write("<script>alert(""请填写回复内容"&CopyRight&""");location=""javascript:history.back(1)"";</script>")  
			Response.End
		End if
		If Len(request.form("Content"))>RsAdminConfigObj("MaxContent")+3000 then
			Response.Write("<script>alert(""内容不能超过"& RsAdminConfigObj("MaxContent")+3000 &"字符"&CopyRight&""");location=""javascript:history.back(1)"";</script>")  
			Response.End
		End if
	  Dim Rs,Sql1
	  Set Rs = server.createobject(G_FS_RS)
	  Sql1 = "select * from FS_GBook where 1=0"
	  Rs.open sql1,conn,1,3
	  Rs.addnew
	  Rs("Content")=Trim(NoCSSHackContent(Request.Form("Content")))
	  Rs("AddTime")=Now()
	  Rs("UserID")=0
	  Rs("FaceNum")=NoCSSHackInput(Replace(request.form("FaceNum"),"'",""))
	  Rs("isQ")=0
	  Rs("isAdmin")=0
	  Rs("Orders")=2
	  Rs("isLock")=0
	  Rs("EditQ")=""
	  Rs("QID")=NoCSSHackInput(Replace(Request.form("QID"),"'",""))
	  Rs.update
	  '更新恢复帖子
	   Conn.execute("Update FS_GBook Set isQ = 1,Qtime="&StrSqlDate&" where id="&Replace(Replace(Request.form("QID"),"'",""),Chr(39),""))
	  Response.Write("<script>alert(""回复成功"&CopyRight&""");location=""SysBookRead.asp?ID="& Replace(request.form("QID"),"'","") &""";</script>") 
	  Response.End
	  Rs.close
	  Set rs=nothing
End if
iF Request("Action")="sLock" then
	Conn.execute("Update FS_GBook Set isLock=1 where id="&Replace(Request("Id"),"'",""))
	  Response.Write("<script>alert(""锁定成功"&CopyRight&""");location=""SysBookRead.asp?ID="& Replace(request("sID"),"'","") &""";</script>") 
	  Response.End
End if
iF Request("Action")="sUnLock" then
	Conn.execute("Update FS_GBook Set isLock=0 where id="&Replace(Request("Id"),"'",""))
	  Response.Write("<script>alert(""解锁成功"&CopyRight&""");location=""SysBookRead.asp?ID="& Replace(request("sID"),"'","") &""";</script>") 
	  Response.End
End if
iF Request("Action")="Top" then
	Conn.execute("Update FS_GBook Set Orders=1 where id="&Replace(Request("Id"),"'",""))
	  Response.Write("<script>alert(""固顶成功"&CopyRight&""");location=""SysBookRead.asp?ID="& Replace(request("sID"),"'","") &""";</script>") 
	  Response.End
End if
iF Request("Action")="UnTop" then
	Conn.execute("Update FS_GBook Set Orders=2 where id="&Replace(Request("Id"),"'",""))
	  Response.Write("<script>alert(""解固成功"&CopyRight&""");location=""SysBookRead.asp?ID="& Replace(request("sID"),"'","") &""";</script>") 
	  Response.End
End if
iF Request("Action")="Del" then
	if Not JudgePopedomTF(Session("Name"),"P070704") then Call ReturnError1()
	Dim GBListObj1
	Set GBListObj1 = Conn.execute("Select ID,UserID From FS_GBook where ID="&Replace(Request("Id"),"'",""))
	If Cint(GBListObj1("UserID"))<>0 Then
		Conn.execute("Update FS_Members Set Point=Point-"&RsAdminConfigObj("QPoint")&" where id="&GBListObj1("UserID"))
	End if
	Conn.execute("Delete From FS_GBook where id="&Replace(Request("Id"),"'",""))
	'扣除会员积分
	If Request("GetAction")="1" then
		Response.Write("<script>alert(""删除成功"&CopyRight&""");location=""SysBook.asp"";</script>") 
	Else
		Response.Write("<script>alert(""删除成功"&CopyRight&""");location=""SysBookRead.asp?id="&Request("sid")&""";</script>") 
	End if 
	Response.End
End if
Dim NewsContent
NewsContent = Replace(Replace(Request.Form("Content"),"""","%22"),"'","%27")
Dim RsModifyObj,ModifySQL
  Set RsModifyObj = server.createobject(G_FS_RS)
  ModifySQL = "select * from FS_GBook where ID="&Replace(Replace(Request("Id"),"'",""),Chr(39),"")
  RsModifyObj.open ModifySQL,conn,1,1
Dim MemberObj
Set MemberObj = Conn.execute("Select Sex,HeadPic,MemName,OICQ,Email From FS_Members Where id="&Replace(Replace(RsModifyObj("Userid"),"'",""),Chr(39),""))
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>FoosunCMS Shop 1.0.0930</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<body scroll=yes topmargin="2" leftmargin="2"> 
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0"> 
	<tr> 
		<td height="398" valign="top"> <div align="left"> 
				<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999"> 
					<tr bgcolor="#EEEEEE"> 
						<td height="26" colspan="5" valign="middle"> <table height="22" border="0" cellpadding="0" cellspacing="0"> 
								<tr> 
									<td width=55 align="center" alt="回复留言" onClick="Reply();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">回复留言</td> 
									<td width=2 class="Gray">|</td> 
									<td width=35 id="SaveID1" style="display:none" align="center" alt="保存" onClick="SubmitFun();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td> 
									<td width=2 id="SaveID2" class="Gray" style="display:none">|</td>
									<td width=55 align="center" alt="留言搜索" onClick="SearchLyPage();" onMouseMove="BtnMouseOver(this);//ShowAddMenu();" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">留言搜索</td> 
									<td width=2 class="Gray">|</td>
									<td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);//ShowAddMenu();" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td> 
								</tr> 
							</table></td> 
					</tr> 
				</table> 
				<table width="100%" border="0" cellspacing="0" cellpadding="5" id="SearchPage" style="display:none;"> 
					<tr> 
						<td width="9%"><a href="SysBook.asp">留言搜索</a><a href="SysBook.asp?Action=UnQ"></a></td> 
						<form name="form1" method="post" action="SysBook.asp"> 
							<td width="91%"><input name="Keyword" type="text" id="Keyword"> 
								<input type="submit" name="Submit2" value="搜索"> </td> 
						</form> 
					</tr> 
				</table> 
				<TABLE cellSpacing=0 cellPadding=0 width="100%" align=center 
                  border=0> 
					<TBODY> 
						<tr> 
							<td height="2"></td> 
						</tr> 
						<TR> 
							<TD width="100%" height="159" valign="top">
								<table width="100%" height="114" border="0" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC"> 
									<tr bgcolor="#FFFFFF"> 
										<td width="17%" valign="top" bgcolor="#EFEFEF"> <%
					If RsModifyObj("UserID")<>0  then
					%> 
											<table width="100%" border="0" cellspacing="0" cellpadding="0"> 
												<tr> 
													<td width="71%"> <strong> <a href=../ReadUser.asp?UserName=<% = MemberObj("MemName")%>> 
														<% = MemberObj("MemName")%> 
														</a> </strong></td> 
													<td width="29%"> <%
						  If MemberObj("Sex") =0 then
						  %> 
														<img src="../../../<%=UserDir%>/GBook/Images/Male.gif" alt="帅哥哦" width="23" height="21"> 
														<%Else%> 
														<img src="../../../<%=UserDir%>/GBook/Images/FeMale.gif" alt="美女哦" width="23" height="21"> 
														<%End if%> </td> 
												</tr> 
											</table> 
											<div align="center"> 
												<hr size="1" noshade color="#CCCCCC"> 
												<%If Len(MemberObj("HeadPic"))>5 then%> 
												<img src=../../../<%=UserDir%>/<% = MemberObj("HeadPic")%>> 
												<%Else%> 
												<table width="0" border="0" cellpadding="0" cellspacing="0" bgcolor="#F0F0F0"> 
													<tr> 
														<td bgcolor="#FFFFFF"><img src="../../../<%=UserDir%>/images/noHeadPic.jpg" width="50" height="50" border="0"></td> 
													</tr> 
												</table> 
												<%End if%> 
												<%Else%> 
												<strong><font color="#990033">管理员</font></strong> 
												<%End if%> 
											</div> 
											<br> 
											<% = RsModifyObj("addtime")%> </td> 
										<td width="83%" valign="top"> <table width="100%" border="0" cellspacing="0" cellpadding="3"> 
												<tr> 
													<td colspan="2" valign="top"> <%If RsModifyObj("UserID")<>0  then%> 
														<table width="100%" border="0" cellspacing="0" cellpadding="3"> 
															<tr> 
																<td width="86"> <div align="center"> 
																		<%
						if Trim(MemberObj("OICQ"))<>"" then
							Dim sOICQ
						    sOICQ ="<a target=blank href=http://wpa.qq.com/msgrd?V=1&Uin="& MemberObj("OICQ") &"&Site="& RsAdminConfigObj("SiteName") &"&Menu=yes><img border=""0"" SRC=http://wpa.qq.com/pa?p=1:"& MemberObj("OICQ") &":8 alt=""点击这里给"& MemberObj("OICQ") &"发消息""></a>"
							Response.Write sOICQ
						Else
							Response.Write("没有QQ")
						End if
						%> 
																	</div></td> 
																
                                <td width="48"><a href="../../../<%=UserDir%>/ReadUser.asp?UserName=<% = MemberObj("MemName")%>"><img src="../../../<%=UserDir%>/GBook/Images/profile.gif" alt="查看信息" width="45" height="18" border="0"></a></td>
                                <td width="45"><a href="mailto:<%=MemberObj("Email")%>"><img src="../../../<%=UserDir%>/GBook/Images/email.gif" width="45" height="18" border="0"></a></td>
                                <td width="45">&nbsp;</td> 
																
                                <td width="113">&nbsp;</td> 
																<td width="113">&nbsp;</td> 
																<td width="113">&nbsp;</td> 
																<td width="113">&nbsp;</td> 
																<td width="113">&nbsp;</td> 
																<td width="113"><div align="right">楼主</div></td> 
															</tr> 
															<tr bgcolor="#D0D0D0"> 
																<td height="1" colspan="10"></td> 
															</tr> 
														</table> 
														<%End if%> </td> 
												</tr> 
												<tr> 
													<td width="4%" valign="top"><img src="../../../<%=UserDir%>/GBook/Images/face<% = RsModifyObj("FaceNum")%>.gif" width="22" height="22"></td> 
													<td width="96%" valign="bottom"><strong> 
														<% = RsModifyObj("Title")%> 
														</strong></td> 
												</tr> 
												<tr> 
													<td height="29" valign="top">&nbsp;</td> 
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
															<%If RsModifyObj("Orders")=2 then%> 
															<a href="SysBookRead.asp?ID=<%=RsModifyObj("id")%>&Sid=<%=Request("id")%>&Action=Top" title=固顶>[固顶]</a> 
															<%Else%> 
															<a href="SysBookRead.asp?ID=<%=RsModifyObj("id")%>&Sid=<%=Request("id")%>&Action=UnTop">[解固]</a> 
															<%End if%> 
															<%If RsModifyObj("isLock")=0 then%> 
															<a href="SysBookRead.asp?ID=<%=RsModifyObj("id")%>&Sid=<%=Request("id")%>&Action=sLock" title=锁定后前台用户不能回复>[锁定]</a> 
															<%Else%> 
															<a href="SysBookRead.asp?ID=<%=RsModifyObj("id")%>&Sid=<%=Request("id")%>&Action=sUnLock">[解锁]</a> 
															<%End if%> 
															<a href="SysBookModify.asp?ID=<%=RsModifyObj("id")%>&GetAction=oper&Sid=<%=Request("id")%>">[编辑]</a> <a href="SysBookRead.asp?ID=<%=RsModifyObj("id")%>&Action=Del&GetAction=1">[删除]</a> </div></td> 
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
										<td valign="top" bgcolor="#EFEFEF"> <%
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
													<td width="29%"> <%
						  If QMemberObj("Sex") =0 then
						  %> 
														<img src="../../../<%=UserDir%>/GBook/Images/Male.gif" alt="帅哥哦" width="23" height="21"> 
														<%Else%> 
														<img src="../../../<%=UserDir%>/GBook/Images/FeMale.gif" alt="美女哦" width="23" height="21"> 
														<%End if%> </td> 
												</tr> 
											</table> 
											<div align="center"> 
												<hr size="1" noshade color="#CCCCCC"> 
												<%If Len(QMemberObj("HeadPic"))>5 then%> 
												<img src=../../../<%=UserDir%>/<% = QMemberObj("HeadPic")%>> 
												<%Else%> 
												<table width="0" border="0" cellpadding="0" cellspacing="0" bgcolor="#F0F0F0"> 
													<tr> 
														<td bgcolor="#FFFFFF"><img src="../../../<%=UserDir%>/images/noHeadPic.jpg" width="50" height="50" border="0"></td> 
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
												<%=RsQModifyObj("Addtime")%> <br> 
											</div></td> 
										<td valign="top"> <%If RsQModifyObj("UserID")<>0  then%> 
											<table width="100%" border="0" cellspacing="0" cellpadding="3"> 
												<tr> 
													<td width="10%"> <div align="center"> 
															<%
						if Trim(QMemberObj("OICQ"))<>"" then
							Dim sOICQ1
						    sOICQ1 ="<a target=blank href=http://wpa.qq.com/msgrd?V=1&Uin="& QMemberObj("OICQ") &"&Site="& RsAdminConfigObj("SiteName") &"&Menu=yes><img border=""0"" SRC=http://wpa.qq.com/pa?p=1:"& QMemberObj("OICQ") &":8 alt=""点击这里给"& QMemberObj("OICQ") &"发消息""></a>"
							Response.Write sOICQ1
						Else
							Response.Write("没有QQ")
						End if
						%> 
														</div></td> 
													
                          <td width="7%"><a href="../../../<%=UserDir%>/ReadUser.asp?UserName=<% = QMemberObj("MemName")%>"><img src="../../../<%=UserDir%>/GBook/Images/profile.gif" alt="查看信息" width="45" height="18" border="0"></a></td>
                          <td width="6%"><a href="mailto:<%=QMemberObj("Email")%>"><img src="../../../<%=UserDir%>/GBook/Images/email.gif" width="45" height="18" border="0"></a></td>
                          <td width="6%">&nbsp;</td> 
													
                          <td width="7%">&nbsp;</td> 
													<td width="16%">&nbsp;</td> 
													<td width="11%">&nbsp;</td> 
													<td width="11%">&nbsp;</td> 
													<td width="6%">&nbsp;</td> 
													<td width="20%"><div align="right"><%=I%>楼</div></td> 
												</tr> 
												<tr bgcolor="#D0D0D0"> 
													<td height="1" colspan="10"></td> 
												</tr> 
											</table> 
											<%Else%> 
											<table width="100%" border="0" cellspacing="0" cellpadding="3"> 
												<tr> 
													<td width="10%"> <div align="center"> </div></td> 
													<td width="7%">&nbsp;</td> 
													<td width="6%">&nbsp;</td> 
													<td width="6%">&nbsp;</td> 
													<td width="7%">&nbsp;</td> 
													<td width="16%">&nbsp;</td> 
													<td width="11%">&nbsp;</td> 
													<td width="11%">&nbsp;</td> 
													<td width="6%">&nbsp;</td> 
													<td width="20%"><div align="right"><%=I%>楼</div></td> 
												</tr> 
												<tr bgcolor="#D0D0D0"> 
													<td height="1" colspan="10"></td> 
												</tr> 
											</table> 
											<%End if%> 
											<table width="100%" height="107" border="0" cellpadding="0" cellspacing="0"> 
												<tr> 
													<td width="4%" height="30" valign="top"><img src="../../../<%=UserDir%>/GBook/Images/face<% = RsQModifyObj("FaceNum")%>.gif" width="22" height="22"></td> 
													<td width="96%" valign="top"> <table width="100%" border="0" cellspacing="0" cellpadding="0"> 
															<tr> 
																<td height="8"></td> 
															</tr> 
														</table> 
														<% = RsQModifyObj("Content")&RsQModifyObj("EditQ")%> </td> 
												</tr> 
												<tr> 
													<td valign="top">&nbsp;</td> 
													<td valign="bottom"> <div align="right"> <a href="SysBookModify.asp?ID=<%=RsQModifyObj("id")%>&GetAction=oper&Sid=<%=Request("id")%>">[编辑]</a> <a href="SysBookRead.asp?ID=<%=RsQModifyObj("id")%>&Action=Del&Sid=<%=Request("id")%>&GetAction=2">[删除]</a> </div></td> 
												</tr> 
											</table></td> 
									</tr> 
									<%
					     RsQModifyObj.MoveNext
					 Next
					%> 
									<tr bgcolor="#FFFFFF"> 
										<td valign="top" bgcolor="#EFEFEF">&nbsp;</td> 
										<td valign="top"> <%
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
					%> </td> 
									</tr> 
									<%End if%> 
								</table></TD> 
						</TR> 
					</TBODY> 
				</TABLE>  
				<%
If Request("QAction")="Q" Then
	if Not JudgePopedomTF(Session("Name"),"P070705") then Call ReturnError1()
	IsShowSave="1"
	Call QuickQ()
Else
	IsShowSave="0"
End if
Sub QuickQ()
%> 
				<a name="QU"></a> 
				<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
				<tr>
				<td height=5 bgcolor="#FFFFFF" colspan="2"></td></tr>
					<form action="SysBookRead.asp" method="POST" name="NewsForm"> 
						<tr bgcolor="#FFFFFF"> 
							<td height="19" colspan="2" bgcolor="#EFEFEF"> <div align="left"><strong>快速回复帖子<a name="B"></a></strong></div></td> 
						</tr> 
						<tr bgcolor="#FFFFFF"> 
							<td width="16%" bgcolor="#F3F3F3"> <div align="right">表情：</div></td> 
							<td width="84%"> <table width="100%" border="0" cellspacing="0" cellpadding="0"> 
									<tr> 
										<td> <input name="FaceNum" type="radio" value="1" checked> 
											<img src="../../../<%=UserDir%>/GBook/Images/face1.gif" width="22" height="22"> </td> 
										<td> <input type="radio" name="FaceNum" value="2"> 
											<img src="../../../<%=UserDir%>/GBook/Images/face2.gif" width="22" height="22"> </td> 
										<td> <input type="radio" name="FaceNum" value="3"> 
											<img src="../../../<%=UserDir%>/GBook/Images/face3.gif" width="22" height="22"> </td> 
										<td> <input type="radio" name="FaceNum" value="4"> 
											<img src="../../../<%=UserDir%>/GBook/Images/face4.gif" width="22" height="22"> </td> 
										<td> <input type="radio" name="FaceNum" value="5"> 
											<img src="../../../<%=UserDir%>/GBook/Images/face5.gif" width="22" height="22"> </td> 
										<td> <input type="radio" name="FaceNum" value="6"> 
											<img src="../../../<%=UserDir%>/GBook/Images/face6.gif" width="22" height="22"></td> 
										<td> <input type="radio" name="FaceNum" value="7"> 
											<img src="../../../<%=UserDir%>/GBook/Images/face7.gif" width="22" height="22"> </td> 
										<td> <input type="radio" name="FaceNum" value="8"> 
											<img src="../../../<%=UserDir%>/GBook/Images/face8.gif" width="22" height="22"> </td> 
										<td> <input type="radio" name="FaceNum" value="9"> 
											<img src="../../../<%=UserDir%>/GBook/Images/face9.gif" width="22" height="22"></td> 
									</tr> 
									<tr> 
										<td> <input type="radio" name="FaceNum" value="10"> 
											<img src="../../../<%=UserDir%>/GBook/Images/face10.gif" width="22" height="22"></td> 
										<td> <input type="radio" name="FaceNum" value="11"> 
											<img src="../../../<%=UserDir%>/GBook/Images/face11.gif" width="22" height="22"> </td> 
										<td> <input type="radio" name="FaceNum" value="12"> 
											<img src="../../../<%=UserDir%>/GBook/Images/face12.gif" width="22" height="22"></td> 
										<td> <input type="radio" name="FaceNum" value="13"> 
											<img src="../../../<%=UserDir%>/GBook/Images/face13.gif" width="22" height="22"> </td> 
										<td> <input type="radio" name="FaceNum" value="14"> 
											<img src="../../../<%=UserDir%>/GBook/Images/face14.gif" width="22" height="22"></td> 
										<td> <input type="radio" name="FaceNum" value="15"> 
											<img src="../../../<%=UserDir%>/GBook/Images/face15.gif" width="22" height="22"></td> 
										<td> <input type="radio" name="FaceNum" value="16"> 
											<img src="../../../<%=UserDir%>/GBook/Images/face16.gif" width="22" height="22"></td> 
										<td> <input type="radio" name="FaceNum" value="17"> 
											<img src="../../../<%=UserDir%>/GBook/Images/face17.gif" width="22" height="22"> </td> 
										<td> <input type="radio" name="FaceNum" value="18"> 
											<img src="../../../<%=UserDir%>/GBook/Images/face18.gif" width="22" height="22"> 
											<input name="Content" type="hidden" id="Content" value="<% = NewsContent %>"> 
											<input name="QID" type="hidden" id="QID" value="<% = Request("ID") %>"> 
											<input name="action" type="hidden" id="action" value="add"> </td> 
									</tr> 
								</table></td> 
						</tr> 
						<tr bgcolor="#FFFFFF"> 
							<td colspan="2"> <div align="right"></div> 
								<iframe id='NewsContent' src='../../../<%=UserDir%>/Editer/BookQNewsEditer.asp' frameborder=0 scrolling=no width='100%' height='200'></iframe></td> 
						</tr> 
					</form> 
				</table> 
				<%End Sub%> 
			</div></td> 
	</tr> 
</table> 
</body>
</html>
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
function Reply()
{
	location='SysBookRead.asp?id=<%=RsModifyObj("ID")%>&QAction=Q#B'
}
function ShowSave()
{
	if(<%=IsShowSave%>=='1')
	{
		SaveID1.style.display='';
		SaveID2.style.display='';
	}
}
ShowSave();
function SearchLyPage()
{
	SearchPage.style.display='';
}
</script>
