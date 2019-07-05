<!--#include file="Inc/Cls_DB.asp" -->
<!--#include file="Inc/Const.asp" -->
<!--#include file="Inc/Md5.asp" -->
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
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="Inc/Function.asp" -->
<%
 if Request("Newsid")="" and  Request("Downloadid")="" Then
		Response.Write("<script>alert(""错误：\n错误的参数,请检查"");location.href=""javascript:history.go(-1)"";</script>")
		Response.End
end if
Dim TempRsNewsObj,TempFlag,Downloadid,Newsid
TempFlag = true
Newsid=Replace(Replace(Trim(Request("Newsid")),"'",""),Chr(39),"")
Downloadid=Replace(Replace(Trim(Request("Downloadid")),"'",""),Chr(39),"")
if Newsid <> "" Then
	Set TempRsNewsObj = Conn.Execute("Select ReviewTF from FS_News where Newsid='" & Newsid & "'")
	if Not TempRsNewsObj.Eof then
		if cint(TempRsNewsObj("ReviewTF")) = 0 then
			TempFlag = False
		end if
	else
		TempFlag = False
	end if
	if TempFlag = False then
		Response.Write("<script>alert(""错误：\n此新闻不允许评论"");location.href=""javascript:history.go(-1)"";</script>")
		Response.End
	end if
elseif Downloadid <> "" Then
	Set TempRsNewsObj = Conn.Execute("Select ReviewTF from FS_Download where Downloadid='" & Downloadid & "'")
	if Not TempRsNewsObj.Eof then
		if cint(TempRsNewsObj("ReviewTF")) = 0 then
			TempFlag = False
		end if
	else
		TempFlag = False
	end if
	if TempFlag = False then
		Response.Write("<script>alert(""错误：\n此下载不允许评论"");location.href=""javascript:history.go(-1)"";</script>")
		Response.End
	end if
end if
if request.Form("action")="add" then
	if  request.Form("NoName")="" then
		if request.Form("MemName")="" then
			Response.Write("<script>alert(""错误：\n请填写您的用户名，匿名用户请选择！"");location.href=""javascript:history.go(-1)"";</script>")
			Response.End
		end if 
		if request.Form("Password")="" then
			Response.Write("<script>alert(""错误：\n请填写您的密码！"");location.href=""javascript:history.go(-1)"";</script>")
			Response.End
		end if
		set Rs = server.CreateObject(G_FS_RS)
		Sql = "select * from FS_Members where MemName='" &Replace(Replace(trim(request("MemName")),"'",""),Chr(39),"")&"' and Password='"&MD5(Replace(Replace(trim(request("Password")),"'","''"),Chr(39),""),16)&"'"
		Rs.Open Sql,Conn,1,3
		if rs.eof then
			Response.Write("<script>alert(""错误：\n没有这个用户,或者密码错误，请重新填写！"");location.href=""javascript:history.go(-1)"";</script>")
			Response.End
	    Else
			Session("MemName") = Rs("MemName")
			Session("MemPassWord") = Rs("Password")
			Session("MemID") = Rs("ID")
			Session("GroupID") = Rs("GroupID")
			Session("Point") = Rs("Point")
			Response.Cookies("Foosun")("MemName") = Rs("MemName")
			Response.Cookies("Foosun")("MemPassword") = Rs("Password")
			Response.Cookies("Foosun")("MemID") = Rs("ID")
			Response.Cookies("Foosun")("GroupID") = Rs("GroupID")
			Response.Cookies("Foosun")("Point") = Rs("Point")
			Session("RePassWord") = Replace(Replace(trim(request("Password")),"'","''"),Chr(39),"")
			Dim Rscon
			set Rscon= conn.execute("select NumberContPoint,NumberLoginPoint from FS_Config")
			conn.execute("update FS_members set LoginNum=LoginNum+1,Point=Point+"&clng(Rscon("NumberLoginPoint"))&",LastLoginIP='"&trim(Request.ServerVariables("Remote_ADDR"))&"',LastLoginTime='" & date() & "' where MemName='"&Rs("MemName")&"'")'用户登陆一次，积分+1分
			Rscon.close
			set Rscon=nothing
		end if 
	End if
		if request.Form("RevContent")="" then
			Response.Write("<script>alert(""错误：\n请输入评论内容！"");location.href=""javascript:history.go(-1)"";</script>")
			Response.End
		end if
		if Len(request.Form("RevContent"))>300 then
			Response.Write("<script>alert(""错误：\n评论不能大于300个字符！"");location.href=""javascript:history.go(-1)"";</script>")
			Response.End
		end if
		Dim Rscon1
		Set Rscon1= conn.execute("select ReviewShow from FS_Config")
		set Rs1 = server.CreateObject(G_FS_RS)
		Sql1 = "select * from FS_Review where 1=0"
		Rs1.Open Sql1,Conn,1,3
		Rs1.addnew
		if Request.Form("NoName")="" then
			Rs1("UserID")=Replace(request("MemName"),"'","''")
		else
			Rs1("UserID")="匿名"
		end if
		if Newsid <> "" Then
			Rs1("NewsID")=Replace(Request.form("NewsID"),"'","''")
			Rs1("Types") = 1
		elseif Downloadid <> "" Then
			Rs1("NewsID")=Replace(Request.form("DownloadID"),"'","''")
			Rs1("Types") = 2
		End if
		Rs1("Content")=Request.form("RevContent")
		If Rscon1("ReviewShow")=0 then
			Rs1("Audit") = 1
		Else
			Rs1("Audit") = 0
		End if
		Rs1("IP")=Request.ServerVariables("Remote_Addr")
		Rs1("AddTime")=now()
		Rs1("Isv")=1
		Rs1.update
		if Newsid <> "" Then
			Response.Redirect("NewsReview.asp?Newsid="& Newsid&"")
		elseif Downloadid <> "" Then
			Response.Redirect("NewsReview.asp?Downloadid="&Downloadid&"")
		End if
		response.end 
end if
strpage=request.querystring("page")
		if len(strpage)=0 then
		strpage="1"
		end if
Set RsConfigObj = Conn.execute("select SiteName,Domain,UseDatePath From FS_Config")
set Rs = server.CreateObject(G_FS_RS)
if Newsid <> "" Then
	Sql = "select * from FS_Review where Newsid='" &Newsid &"' and  Types = 1  and isv=1 and Audit=1 order by ID desc"
elseif Downloadid <> "" Then
	Sql = "select * from FS_Review where Newsid='" &Downloadid&"' and Types = 2 and isv=1  and Audit=1 order by ID desc"
end if
Rs.Open Sql,Conn,1,1
%>
<html>
<title><% = RsConfigObj("SiteName") %>_____用户评论</title>
<link href="CSS/FS_css.css" rel="stylesheet">
<body bgcolor="#FFFFFF">
<table width="95%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#D7D7D7" class="tabbgcolor">
  <tr class="tabbgcolorliWhite"> 
    <td colspan="2" bgcolor="#FFFFFF"> <TABLE width="100%" border=0 cellpadding="5" cellspacing="0">
        <TBODY>
          <TR> 
            <TD width=26><IMG 
                              src="<%=UserDir%>/images/Favorite.OnArrow.gif" border=0></TD>
            <TD bgcolor="#FFFFFF" 
class=f4>新闻/下载评论</TD>
          </TR>
        </TBODY>
      </TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" 
border=0>
        <TBODY>
          <TR> 
            <TD bgColor=#ff6633 height=4><IMG height=1 src="" 
                              width=1></TD>
          </TR>
        </TBODY>
      </TABLE></td>
  </tr>
  <tr class="tabbgcolorliWhite"> 
    <td width="78%" colspan="2" bgcolor="#FFFFFF"> 
      <%
if Rs.eof and Rs.bof then
	Response.write "<p align='center'> 未找到评论</p>"
	Else
	rs.pagesize=20
	rs.absolutepage=cint(strpage)
	select_count=rs.recordcount
	select_pagecount=rs.pagecount
	%> <table width="100%" border="0" cellspacing="0" cellpadding="6">
        <%
		 for i=1 to rs.pagesize
		if rs.eof then
		exit for
		end if
		%>
        <tr> 
          <td height="17" colspan="2" bgcolor="#F5F5F5">·来自：<font color="#0000FF"><%=Replace(rs("IP"),Mid(rs("IP"),InstrRev(rs("IP"),".")+1),"**")%></font>的<font color="#FF0000"> 
            <%
		  set Rs2 = server.CreateObject(G_FS_RS)
		 Sql2 = "select * from FS_Members where MemName='" &Replace(rs("Userid"),"'","''")&"'"
		 Rs2.Open Sql2,Conn,1,3
		  if rs("Userid")="匿名" or rs("Userid")="" then
		     Members="匿名用户"
		  else
		     Members= "<a href="& UserDir &"/ReadUser.asp?UserName="&rs2("MemName")&" target=_blank>"&rs("Userid")&"</a>"
		  end if		  
		  %>
            </font><strong> 
            <% = Members%>
            </strong>于<%=rs("AddTime")%>对 
            <%
			set Rs1 = server.CreateObject(G_FS_RS)
			if Newsid <> "" Then
				Sql1 = "select * from FS_News where NewsId='" &Replace(request("NewsId"),"'","''")&"'"
				Rs1.Open Sql1,Conn,1,1
				Dim NewsPath
				If RsConfigObj("UseDatePath")=1 then
					NewsPath = Rs1("Path")
				Else
					NewsPath = ""
				End if
				Dim RsClassObj
				Set RsClassObj = Conn.execute("Select ClassID,ClassEname,SaveFilePath From FS_NewsClass Where ClassID='"& Replace(Replace(Rs1("ClassID"),"'",""),Chr(39),"")&"'")
				%> <a href=<%=RsConfigObj("Domain")&RsClassObj("SaveFilePath")&"/"&RsClassObj("ClassEname")&NewsPath&"/"&Rs1("FileName")&"."&Rs1("FileExtName")&""%> target="_blank"><font color="#FF0000"><%=rs1("Title")%></font></a> <%
				rs1.close
				set rs1=nothing
				Set  RsClassObj= nothing
			elseif Downloadid <> "" Then
				Sql1 = "select * from FS_Download where DownloadId='" &Replace(request("DownloadId"),"'","''")&"'"
				Rs1.Open Sql1,Conn,1,1
				Set RsClassObj = Conn.execute("Select ClassID,ClassEname,SaveFilePath From FS_NewsClass Where ClassID='"& Replace(Replace(Rs1("ClassID"),"'",""),Chr(39),"")&"'")
				%> <a href=<%=RsConfigObj("Domain")&RsClassObj("SaveFilePath")&"/"&RsClassObj("ClassEname")&"/"&Rs1("FileName")&"."&Rs1("FileExtName")&""%> target="_blank"><font color="#FF0000"><%=rs1("Name")%></font></a> <%
				rs1.close
				set rs1=nothing
				Set RsClassObj = Nothing
			End if
			%>
            发表的的评论：</td>
        </tr>
        <tr> 
          <td height="39" colspan="2" valign="top"> <%
		if conn.execute("select ReviewShow from FS_Config")(0) = 1 then
			if RS("Audit") = 1 then
			  Response.Write(rs("Content"))
			else
			  Response.Write("<font color=""red"">管理员还没有审核此评论,暂时不显示。</font>")
			end if
		else
			  Response.Write(rs("Content"))
		end if
		  %> </td>
        </tr>
        <%
	  rs.movenext
	 next
	%>
      </table>
      <%
	   response.write"&nbsp;&nbsp;共<b>"& select_pagecount &"</b>页<b>" & select_count &"</b>条记录，本页是第<b>"& strpage &"</b>页。"
		if int(strpage)>1 then
		   response.Write"&nbsp;&nbsp;&nbsp;<a href=?page=1&Newsid="&Request("Newsid")&"&Downloadid="&Request("Downloadid")&">第一页</a>&nbsp;"
		   response.Write"&nbsp;&nbsp;&nbsp;<a href=?page="&cstr(cint(strpage)-1)&"&Newsid="&Request("Newsid")&"&Downloadid="&Request("Downloadid")&">上一页</a>&nbsp;"
		end if
		if int(strpage)<select_pagecount then
			response.Write"&nbsp;&nbsp;&nbsp;<a href=?page="&cstr(cint(strpage)+1)&"&Newsid="&Request("Newsid")&"&Downloadid="&Request("Downloadid")&">下一页</a>"
			response.Write"&nbsp;&nbsp;&nbsp;<a href=?page="& select_pagecount &"&Newsid="&Request("Newsid")&"&Downloadid="&Request("Downloadid")&">最后一页</a>&nbsp;"
		end if
		response.Write"<br>"
end if	   
	   %> <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td> </td>
        </tr>
      </table></td>
  </tr>
  <tr class="tabbgcolorliWhite">
    <td colspan="2" bgcolor="#FFFFFF"><form name="form1" method="post" action="">
        <TABLE cellSpacing=0 cellPadding=0 width="100%" 
border=0>
          <TBODY>
            <TR> 
              <TD bgColor=#ff6633 height=4><IMG height=1 src="" 
                              width=1></TD>
            </TR>
          </TBODY>
        </TABLE>
        <table width="100%" border="0" cellpadding="5" cellspacing="1" class="tabbgcolor">
          <tr bgcolor="#F7F7F7"> 
            <td width="15%"> <div align="right">
                <input name="Newsid" type="hidden" id="Newsid" value="<%=trim(Request("Newsid"))%>">
                <input name="Downloadid" type="hidden" id="Downloadid" value="<%=trim(Request("Downloadid"))%>">
                <input name="action" type="hidden" id="action" value="add">
                会员名称：</div></td>
            <td width="85%"> <input name="MemName" type="text" id="MemName" value="<%=session("MemName")%>"> 
              <input name="NoName" type="checkbox" id="NoName" value="1">
              匿名 <font color="#FF0000">·</font><a href="<%=UserDir%>/sRegister.asp"><font color="#FF0000">注册用户</font></a> 
              <a href="<%=UserDir%>/User_GetPassword.asp">·忘记密码？</a>　·<a href="<%=UserDir%>/User_Comments.asp"><font color="#0000FF">查看我的评论</font></a> 
            </td>
          </tr>
          <tr bgcolor="#F7F7F7"> 
            <td> <div align="right">密码：</div></td>
            <td> <input name="Password" type="password" id="Password" value="<%=Session("RePassWord")%>"> </td>
          </tr>
          <tr bgcolor="#F7F7F7"> 
            <td> <div align="right">评论内容：<br>
                (最多300个字符) </div></td>
            <td> <textarea name="RevContent" cols="60" rows="6" id="RevContent"></textarea></td>
          </tr>
          <tr bgcolor="#F7F7F7"> 
            <td colspan="2" align="center"> <input type="submit" name="Submit" value=" 发 表 "> 
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input type="reset" name="Submit2" value=" 重 填 "></td>
          </tr>
        </table>
      </form><br>
      · 尊重网上道德，遵守《全国人大常委会关于维护互联网安全的决定》和《互联网电子公告服务管理规定》及中华人民共和国其他各项有关法律法规。 <br>
      · 严禁发表危害国家安全、损害国家利益、破坏民族团结、破坏国家宗教政策、破坏社会稳定、侮辱、诽谤、教唆、淫秽等内容的作品 。 <br>
      · 用户需对自己在使用本站服务过程中的行为承担法律责任（直接或间接导致的）。 <br>
      · 本论坛版主有权保留或删除其管辖论坛中的任意内容。 <br>
      · 社区内所有的文章版权归原文作者和本站共同所有，任何人需要转载社区内文章，必须征得原文作者或本站授权。 
      <p>· 本贴提交者发言纯属个人意见，与本网站立场无关。 <br>
      </p></td>
  </tr>
</table>
</body></html>
<%
Set RsConfigObj = Nothing
%>