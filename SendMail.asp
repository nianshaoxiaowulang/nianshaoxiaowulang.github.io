<!--#include file="Inc/Cls_DB.asp" -->
<!--#include file="Inc/Const.asp" -->
<%
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
'==============================================================================
'������ƣ���Ѷ��վ��Ϣ����ϵͳ
'��ǰ�汾��Foosun Content Manager System(FoosunCMS V3.1.0930)
'���¸��£�2005.10
'==============================================================================
'Copyright (C) 2002-2004 Foosun.Net  All rights reserved.
'��ҵע����ϵ��028-85098980-601,��Ŀ������028-85098980-606��609,�ͻ�֧�֣�608
'��Ʒ��ѯQQ��394226379,159410,125114015
'����֧��QQ��315485710,66252421 
'��Ŀ����QQ��415637671��655071
'���򿪷����Ĵ���Ѷ�Ƽ���չ���޹�˾(Foosun Inc.)
'Email:service@Foosun.cn
'MSN��skoolls@hotmail.com
'��̳֧�֣���Ѷ������̳(http://bbs.foosun.net)
'�ٷ���վ��www.Foosun.cn  ��ʾվ�㣺test.cooin.com 
'��վͨϵ��(���ܿ��ٽ�վϵ��)��www.ewebs.cn
'==============================================================================
'��Ѱ汾���ڳ�����ҳ������Ȩ��Ϣ�������ϱ�վLOGO��������
'��Ѷ��˾�����˳���ķ���׷��Ȩ��
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
         Response.Write("<script>alert(""û�����Ȩ�ޣ�ԭ�����Ѿ�������\n����ϵͳ����Ա��ϵ"&CopyRight&""");location=""javascript:history.back()"";</script>")  
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

function SendMail(SMTPServer,sender,mailto,subject,mailbody)'�����һ���������ڳ����п���ֱ�ӵ��á�
dim JMail
Set JMail = Server.CreateObject("JMail.SMTPMail") 
JMail.LazySend = true '��������ǽ��ʼ����뻺��ֱ������Ϊֹ���Ĳ����κεĴ�����Ϣ
JMail.Charset = "gb2312" '�趨�ʼ����ַ�����Ĭ��Ϊ"US-ASCII" һ��������"gb2312" 
JMail.ContentType = "text/html" '����ʼ���ͷ���ã� Ĭ��Ϊ "text/plain" �����ó�����Ҫ���κ������ '���뷢��HTML��Ϣ,�����ͷ�ļ�Ϊ "text/html"
JMail.ServerAddress =SMTPServer 'SERVER�ĵ�ַ�������кܶ��SERVER��ַ����ɸ��˿ں�
JMail.Sender = sender'�ʼĵĵ�ַ
JMail.Subject = subject'�ʼ��ı��⡣ 
JMail.AddRecipient mailto'����һ���ռ���
JMail.Body = mailbody 'UBBCode(htmlencode(MSG))E-Mail������
JMail.Priority = 3'�ʼ������ȼ������Է�Χ��1��5��Խ������ȼ�Լ�ߣ����磬5��ߣ�1���,һ������
JMail.AddHeader "Originating-IP", Request.ServerVariables("REMOTE_ADDR")'addheader����һ��ͷ��Ϣ
'remote_addr��������Ļ�����IP��ַ 
'JMail.Execute'ִ���ʼ����͵�SERVER 
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
	Response.write"<script>alert(""����Ĳ�����"");location.href=""javascript:history.back()"";</script>"
    Response.end
end if

	sql="Select * from FS_News where Newsid='"&Newsid&"'"
	set rs=server.createobject(G_FS_RS)
	rs.open sql,conn,1,1
	if rs.bof and rs.eof then
		Response.write"<script>alert(""�Ҳ������ţ�"");location.href=""javascript:history.back()"";</script>"
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
<title>���͵����ʼ�</title>
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
            <TD class=f4>���͵����ʼ�</TD>
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
      <td height="22" colspan=2 align=center valign=middle class="title"> <b>�����ĸ��ߺ���</b></td>
    </tr>
    <tr class="tdbg"> 
      <td width="120" align="right"><strong>������������</strong></td>
      <td><input name="MailtoName" type="text" id="MailtoName" size="60" maxlength="20"> 
        <font color="#FF0000">*</font></td>
    </tr>
    <tr class="tdbg"> 
      <td width="120" align="right"><strong>������Email��ַ��</strong></td>
      <td><input name="MailToAddress" type=text id="MailToAddress" size="60" maxlength="100"> 
        <font color="#FF0000">*</font></td>
    </tr>
    <tr class="tdbg"> 
      <td height="20" align="right"><strong>���������</strong></td>
      <td height="20"> <input name="Username" type=text id="Username" value="<% = Session("sName")%>" size="60" maxlength="100"></td>
    </tr>
    <tr class="tdbg"> 
      <td height="20" align="right"><strong>���Email��ַ��</strong></td>
      <td height="20"><input name="Useremail" type=text id="Useremail" value="<% =Session("email")%>" size="60" maxlength="100"></td>
    </tr>
    <tr class="tdbg"> 
      <td width="120" height="60" align="right"><strong>������Ϣ��</strong></td>
      <td height="60">���ű��⣺<font color="#FF0000"><strong><%= rs("Title") %></strong></font><br>
        �������ߣ�<%= rs("Author") %> <br>
        ����ʱ�䣺<%= rs("Adddate") %> </td>
    </tr>
    <tr class="tdbg"> 
      <td colspan=2 align=center><input name="Action" type="hidden" id="Action" value="MailToFriend"> 
        <input name="filename" type="hidden" id="Newsid" value="<%=request("Newsid")%>"> 
        <input type=submit value=" �� �� " name="Submit" <% If ObjInstalled=false Then response.write "disabled" end if%>> 
      </td>
    </tr>
    <%
If ObjInstalled=false Then
	Response.Write "<tr><td height='40' colspan='2'><b><font color=red>�Բ�����Ϊ��������֧�� JMail���! ���Բ���ʹ�ñ����ܡ�</font></b></td></tr>"
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
		Response.write"<script>alert(""�����˲���Ϊ�գ�"");location.href=""javascript:history.back()"";</script>"
        Response.end
	end if
	if IsValidEmail(MailToAddress)=false then
   		Response.write"<script>alert(""EMAIL��ַ����"");location.href=""javascript:history.back()"";</script>"
        Response.end
	end if
				
	call GetMailInfo()
	
	call SendMail(RsConfig("MailServer"),RsConfig("Sitename"),request.Form("MailToAddress"),subject,mailbody)
	if err then '���
		response.Write("����ʧ��,"&err.description&"")
		response.end
	err.clear
	else
		response.Write("���ͳɹ�")
		response.end
	end if

end sub

sub GetMailInfo()
	Subject="��������"&request.Form("Username")&"��" & RsConfig("SiteName") & "������������������"

	mailbody=mailbody &"<style>A:visited {	TEXT-DECORATION: none	}"
	mailbody=mailbody &"A:active  {	TEXT-DECORATION: none	}"
	mailbody=mailbody &"A:hover   {	TEXT-DECORATION: underline overline	}"
	mailbody=mailbody &"A:link 	  {	text-decoration: none;}"
	mailbody=mailbody &"A:visited {	text-decoration: none;}"
	mailbody=mailbody &"A:active  {	TEXT-DECORATION: none;}"
	mailbody=mailbody &"A:hover   {	TEXT-DECORATION: underline overline}"
	mailbody=mailbody &"BODY   {	FONT-FAMILY: ����; FONT-SIZE: 9pt;}"
	mailbody=mailbody &"TD	   {	FONT-FAMILY: ����; FONT-SIZE: 9pt	}</style>"

	mailbody=mailbody &"<TABLE border=0 width='95%' align=center><TBODY><TR>"
	mailbody=mailbody &"<TD valign=middle align=top>"
	mailbody=mailbody &"--&nbsp;&nbsp;���ߣ�"&rs("Author")&"<br>"
	mailbody=mailbody &"--&nbsp;&nbsp;����ʱ�䣺"&rs("Adddate")&"<br><br>"
	mailbody=mailbody &"--&nbsp;&nbsp;"&rs("title")&"<br>"
	mailbody=mailbody &""&rs("content")&""
	mailbody=mailbody &"</TD></TR></TBODY></TABLE>"

	mailbody=mailbody &"<center><a href='" & RsConfig("DoMain") & "'>" & RsConfig("SiteName") & ",�����ʼ�"&request.Form("Useremail")&"</a>"
end sub
%>
