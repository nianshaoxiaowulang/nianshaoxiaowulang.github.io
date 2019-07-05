<!--#include file="Inc/Cls_DB.asp" -->
<!--#include file="Inc/Const.asp" -->
<%
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
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
Dim LoginUrl,confimsn
Set confimsn=conn.execute("select domain,Copyright from FS_config")
Dim MemName,PassWordm,MemID
MemName =session("MemName")
PassWordm = Session("MemPassword")
MemID = Session("MemID")
Dim RsMemObj
set RsMemObj = Server.CreateObject (G_FS_RS)
RsMemObj.Source="select * from FS_Members where MemName='"& MemName &"' and password='"&PassWordm&"'"
RsMemObj.Open RsMemObj.Source,Conn,1,1
if not RsMemObj.EOF then
      if cint(RsMemObj("Lock"))=1 then
         Response.Write("<script>alert(""没有浏览权限，原因：您已经被锁定\n请与系统管理员联系"&CopyRight&""");location=""javascript:history.back()"";</script>")  
         Response.End
      end if
else
   RsMemObj.Close
   set RsMemObj = nothing
   Dim TopLocation
   TopLocation = ""&confimsn("domain")&"/"&UserDir&"/Login.asp"
   Response.Write("<script>top.location='" & TopLocation & "'</script>")
   Response.end
end if

Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function

Set RsConfig=conn.execute("Select * from FS_Config")

function SendMail(SMTPServer,sender,mailto,subject,mailbody)'这个是一个函数，在程序中可以直接调用。
dim JMail
Set JMail = Server.CreateObject("JMail.SMTPMail") 
JMail.LazySend = true '这个属性是将邮件放入缓冲直至发出为止，的不到任何的错误信息
JMail.Charset = "gb2312" '设定邮件的字符集，默认为"US-ASCII" 一般中文用"gb2312" 
JMail.ContentType = "text/html" '这个邮件的头设置， 默认为 "text/plain" 能设置成你需要的任何情况。 '你想发送HTML信息,改这个头文件为 "text/html"
JMail.ServerAddress =SMTPServer 'SERVER的地址。可以有很多的SERVER地址，后可跟端口号
JMail.Sender = sender'邮寄的地址
JMail.Subject = subject'邮件的标题。 
JMail.AddRecipient mailto'加入一个收件者
JMail.Body = mailbody 'UBBCode(htmlencode(MSG))E-Mail的主体
JMail.Priority = 3'邮件的优先级，可以范围从1到5。越大的优先级约高，比如，5最高，1最低,一般设置
JMail.AddHeader "Originating-IP", Request.ServerVariables("REMOTE_ADDR")'addheader加入一个头信息
'remote_addr发出请求的机器的IP地址 
'JMail.Execute'执行邮件发送到SERVER 
set jMail=nothing   
end function
'----
function IsValidEmail(email)
	dim names, name, i, c
	IsValidEmail = true
	names = Split(email, "@")
	if UBound(names) <> 1 then
	   IsValidEmail = false
	   exit function
	end if
	for each name in names
		if Len(name) <= 0 then
			IsValidEmail = false
    		exit function
		end if
		for i = 1 to Len(name)
		    c = Lcase(Mid(name, i, 1))
			if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
		       IsValidEmail = false
		       exit function
		     end if
	   next
	   if Left(name, 1) = "." or Right(name, 1) = "." then
    	  IsValidEmail = false
	      exit function
	   end if
	next
	if InStr(names(1), ".") <= 0 then
		IsValidEmail = false
	   exit function
	end if
	i = Len(names(1)) - InStrRev(names(1), ".")
	if i <> 2 and i <> 3 then
	   IsValidEmail = false
	   exit function
	end if
	if InStr(email, "..") > 0 then
	   IsValidEmail = false
	end if
end function

ObjInstalled=IsObjInstalled("JMail.SMTPMail")
Newsid= trim(Replace(request("Newsid"),"'","''"))
Action=trim(request("Action"))
if Newsid="" then
	Response.write"<script>alert(""错误的参数！"");location.href=""javascript:history.back()"";</script>"
    Response.end
end if

	sql="Select * from FS_News where Newsid='"&Newsid&"'"
	set rs=server.createobject(G_FS_RS)
	rs.open sql,conn,1,1
	if rs.bof and rs.eof then
		Response.write"<script>alert(""找不到新闻！"");location.href=""javascript:history.back()"";</script>"
		Response.end
	else
		if Action="MailToFriend" then
			call MailToFriend()
		else
			call main()
		end if
	end if
	rs.close
	set rs=nothing
sub main()
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>发送电子邮件</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<link href="CSS/FS_css.css" rel="stylesheet">

<body>
<table width="95%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
  <tr> 
    <td bgcolor="#FFFFFF">
<TABLE width="100%" border=0 cellpadding="6">
        <TBODY>
          <TR> 
            <TD width=26><IMG src="<%=UserDir%>/images/Favorite.OnArrow.gif" border=0></TD>
            <TD class=f4>发送电子邮件</TD>
          </TR>
        </TBODY>
      </TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY>
          <TR> 
            <TD bgColor=#ff6633 height=4><IMG height=1 src="" width=1></TD>
          </TR>
        </TBODY>
      </TABLE></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF">
<form name="form1" method="post" action="">
        <table cellpadding=6 cellspacing=1 border=0 width=90% class="border" align=center>
          <tr> 
      <td height="22" colspan=2 align=center valign=middle class="title"> <b>将本文告诉好友</b></td>
    </tr>
    <tr class="tdbg"> 
      <td width="120" align="right"><strong>收信人姓名：</strong></td>
      <td><input name="MailtoName" type="text" id="MailtoName" size="60" maxlength="20"> 
        <font color="#FF0000">*</font></td>
    </tr>
    <tr class="tdbg"> 
      <td width="120" align="right"><strong>收信人Email地址：</strong></td>
      <td><input name="MailToAddress" type=text id="MailToAddress" size="60" maxlength="100"> 
        <font color="#FF0000">*</font></td>
    </tr>
    <tr class="tdbg"> 
      <td height="20" align="right"><strong>你的姓名：</strong></td>
      <td height="20"> <input name="Username" type=text id="Username" value="<% = Session("sName")%>" size="60" maxlength="100"></td>
    </tr>
    <tr class="tdbg"> 
      <td height="20" align="right"><strong>你的Email地址：</strong></td>
      <td height="20"><input name="Useremail" type=text id="Useremail" value="<% =Session("email")%>" size="60" maxlength="100"></td>
    </tr>
    <tr class="tdbg"> 
      <td width="120" height="60" align="right"><strong>新闻信息：</strong></td>
      <td height="60">新闻标题：<font color="#FF0000"><strong><%= rs("Title") %></strong></font><br>
        新闻作者：<%= rs("Author") %> <br>
        发布时间：<%= rs("Adddate") %> </td>
    </tr>
    <tr class="tdbg"> 
      <td colspan=2 align=center><input name="Action" type="hidden" id="Action" value="MailToFriend"> 
        <input name="filename" type="hidden" id="Newsid" value="<%=request("Newsid")%>"> 
        <input type=submit value=" 发 送 " name="Submit" <% If ObjInstalled=false Then response.write "disabled" end if%>> 
      </td>
    </tr>
    <%
If ObjInstalled=false Then
	Response.Write "<tr><td height='40' colspan='2'><b><font color=red>对不起，因为服务器不支持 JMail组件! 所以不能使用本功能。</font></b></td></tr>"
End If
%>
  </table>
</form>
    </td>
  </tr>
  <tr>
    <td bgcolor="#F2F2F2"> 
      <div align="center">
        <% = confimsn("Copyright") %>
      </div></td>
  </tr>
</table>
</body>
</html>
<%end sub
sub MailToFriend()
	MailToName=trim(request.form("MailToName"))
	MailToAddress=trim(request.form("MailToAddress"))
	if MailToName="" then
		Response.write"<script>alert(""收信人不能为空！"");location.href=""javascript:history.back()"";</script>"
        Response.end
	end if
	if IsValidEmail(MailToAddress)=false then
   		Response.write"<script>alert(""EMAIL地址有误！"");location.href=""javascript:history.back()"";</script>"
        Response.end
	end if
				
	call GetMailInfo()
	
	call SendMail(RsConfig("MailServer"),RsConfig("Sitename"),request.Form("MailToAddress"),subject,mailbody)
	if err then '检测
		response.Write("发送失败,"&err.description&"")
		response.end
	err.clear
	else
		response.Write("发送成功")
		response.end
	end if

end sub

sub GetMailInfo()
	Subject="您的朋友"&request.Form("Username")&"从" & RsConfig("SiteName") & "给您发来的新闻资料"

	mailbody=mailbody &"<style>A:visited {	TEXT-DECORATION: none	}"
	mailbody=mailbody &"A:active  {	TEXT-DECORATION: none	}"
	mailbody=mailbody &"A:hover   {	TEXT-DECORATION: underline overline	}"
	mailbody=mailbody &"A:link 	  {	text-decoration: none;}"
	mailbody=mailbody &"A:visited {	text-decoration: none;}"
	mailbody=mailbody &"A:active  {	TEXT-DECORATION: none;}"
	mailbody=mailbody &"A:hover   {	TEXT-DECORATION: underline overline}"
	mailbody=mailbody &"BODY   {	FONT-FAMILY: 宋体; FONT-SIZE: 9pt;}"
	mailbody=mailbody &"TD	   {	FONT-FAMILY: 宋体; FONT-SIZE: 9pt	}</style>"

	mailbody=mailbody &"<TABLE border=0 width='95%' align=center><TBODY><TR>"
	mailbody=mailbody &"<TD valign=middle align=top>"
	mailbody=mailbody &"--&nbsp;&nbsp;作者："&rs("Author")&"<br>"
	mailbody=mailbody &"--&nbsp;&nbsp;发布时间："&rs("Adddate")&"<br><br>"
	mailbody=mailbody &"--&nbsp;&nbsp;"&rs("title")&"<br>"
	mailbody=mailbody &""&rs("content")&""
	mailbody=mailbody &"</TD></TR></TBODY></TABLE>"

	mailbody=mailbody &"<center><a href='" & RsConfig("DoMain") & "'>" & RsConfig("SiteName") & ",电子邮件"&request.Form("Useremail")&"</a>"
end sub
%>
