<% Option Explicit %>
<!--#include file="../Inc/Cls_DB.asp" -->
<!--#include file="../Inc/Const.asp" -->
<!--#include file="../Inc/Md5.asp" -->
<%
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
Dim DBC,conn
Set DBC = new databaseclass
Set conn = DBC.openconnection()
Set DBC = nothing

Dim I,RsConfigObj
Set RsConfigObj = Conn.Execute("Select * from FS_Config")
'�����ʼ�
Function SendMail(SMTPServer,sender,loginName,LoginPass,mailto,subject,mailbody)'�����һ���������ڳ����п���ֱ�ӵ��á�
	On error resume next
	Dim JMail
	Set JMail = Server.CreateObject("JMail.Message") 
	JMail.Silent = True
	JMail.Logging = True
	JMail.Charset = "gb2312" '�趨�ʼ����ַ�����Ĭ��Ϊ"US-ASCII" һ��������"gb2312" 
	If Not(LoginName = "" Or LoginPass = "") Then
		JMail.MailServerUserName = LoginName '�����ʼ���������¼��
		JMail.MailServerPassword = LoginPass '��¼����
	End If
	JMail.ContentType = "text/html" '����ʼ���ͷ���ã� Ĭ��Ϊ "text/plain" �����ó�����Ҫ���κ������ '���뷢��HTML��Ϣ,�����ͷ�ļ�Ϊ "text/html"
	JMail.From = sender'�ʼĵĵ�ַ
	JMail.FromName = ""&RsConfigObj("SiteName")&"��վ����Ա"
	JMail.Subject = subject'�ʼ��ı��⡣ 
	JMail.AddRecipient mailto'����һ���ռ���
	JMail.Body = mailbody 'E-Mail������
	JMail.Priority = 1'�ʼ������ȼ������Է�Χ��1��5��Խ������ȼ�Լ�ߣ����磬5��ߣ�1���,һ������
	JMail.AddHeader "Originating-IP", Request.ServerVariables("REMOTE_ADDR")'addheader����һ��ͷ��Ϣ
	if not JMail.Send(SMTPServer) then'ִ���ʼ����͵�SERVER  
		SendMail = false
		Response.Write("�ʼ�����ʧ�ܣ������Ƿ�������֧��JMAIL�������ʹ��jmail4.3���ϰ汾��<br>")
	Else
		SendMain = true
		Response.Write("�ʼ��Ѿ����͵���ע��������У���ע�����<br>")
	End If
	JMail.Close
	Set JMail=nothing   
End Function
%>
<HTML><HEAD>
<TITLE><%=RsConfigObj("SiteName")%> >> ע��ɹ�</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="Css/UserCSS.css" type=text/css  rel=stylesheet>
<meta http-equiv="refresh" content="10;URL=main.asp">
</HEAD>
<BODY leftmargin="0" topmargin="10">
<div align="center">
  <script language="JavaScript" src="top.js" type="text/JavaScript"></script>
</div>
<TABLE cellSpacing=2 width="98%" align=center border=0>
  <TBODY>
  <TR>
    <TD vAlign=top width=160>
      <TABLE cellSpacing=0 cellPadding=0 width=102 border=0>
        <TBODY>
        <TR>
          <TD><IMG height=27 src="images/favorite.left.help.jpg" 
        width=190></TD></TR>
        <TR>
          <TD>
            <TABLE style="BORDER-COLLAPSE: collapse" borderColor=#cccccc 
            cellSpacing=0 cellPadding=0 width="100%" border=1>
              <TBODY>
              <TR>
                <TD vAlign=top>
                  <TABLE class=bgup cellSpacing=0 cellPadding=0 width="100%" 
                  background="" border=0>
                    <TBODY>
                    <TR>
                      <TD align=right>&nbsp;</TD>
                      <TD align=right>&nbsp;</TD></TR>
                    <TR>
                      <TD align=right width="15%" height=30>&nbsp;</TD>
                                
                              <TD width="85%"><span class="f4"><STRONG>��д�ʺ���Ϣ</STRONG></span> 
                              </TD>
                              </TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
            <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
              <TBODY>
              <TR>
                <TD height=10><IMG height=1 src="" 
width=1></TD></TR></TBODY></TABLE>
            <TABLE style="BORDER-COLLAPSE: collapse" borderColor=#cccccc 
            cellSpacing=0 cellPadding=0 width="100%" border=1>
              <TBODY>
              <TR>
                <TD vAlign=top>
                  <TABLE class=bgup cellSpacing=0 cellPadding=0 width="100%" 
                  background="" border=0>
                    <TBODY>
                    <TR>
                      <TD align=right>&nbsp;</TD>
                      <TD align=right>&nbsp;</TD></TR>
                    <TR>
                      <TD align=right width="12%" height=30>&nbsp;</TD>
                      <TD width="88%">
                        <TABLE cellSpacing=0 cellPadding=0 width="100%" 
border=0>
                                  <TBODY>
                                    <TR> 
                                      <TD width=14 height=24><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD> ͬ��ע��Э��</TD>
                                    </TR>
                                    <TR> 
                                      <TD width=14 height=24><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD>��д�ʺ���Ϣ</TD>
                                    </TR>
                                    <TR> 
                                      <TD width=14 height=24><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD>��д��ϵ����</TD>
                                    </TR>
                                    <TR> 
                                      <TD width=14 height=24><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><font color="#FF0000">ע��ɹ�</font></TD>
                                    </TR>
                                    <TR> 
                                      <TD><IMG height=5 src="images/SelfService.aspx" 
                              width=1></TD>
                                      <TD></TD>
                                    </TR>
                                  </TBODY>
                                </TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD>
      <TD vAlign=top> <TABLE cellSpacing=0 cellPadding=0 width="98%" align=center 
                  border=0>
          <TBODY>
            <TR> 
              <TD width="100%"> <TABLE width="100%" border=0>
                  <TBODY>
                    <TR> 
                      <TD width=26><IMG 
                              src="images/Favorite.OnArrow.gif" border=0></TD>
                      <TD 
class=f4>ע��ɹ�</TD>
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
              <form method=POST action="sRegister_Success.asp" name=UserForm1 onSubmit="return checkdata1()">
                <TD width="100%" height="159"> 
                  <div align="center"><br>
                    ��ϲ��!<font color="#FF0000"><strong><%=Session("sName")%></strong></font>������<%=RsConfigObj("SiteName")%>ע��ɹ�<br>
                    <br>
					<%
					if cint(RsConfigObj("isEmail"))=1 then
						Call SendMail(trim(RsConfigObj("MailServer")),trim(RsConfigObj("Email")),trim(RsConfigObj("MailName")),RsConfigObj("MailPass"),Session("email"),"����["&RsConfigObj("SiteName")&"]��ע����Ϣ",Session("sName")&":���ã�<br>��ӭע���Ϊ"&RsConfigObj("SiteName")&"��Ա,�����û�����"& Session("MemName") &"�����룺"& Session("VerPassword") &",��½��ַ��"&RsConfigObj("DoMain")&"/User/Login.asp")
						Response.Write("<br>")
					End if
					%>
                    10��󷵻�<a href="main.asp"><font color="#FF0000">��Ա������ҳ</font></a> 
                  </div></TD></form>
            </TR>
          </TBODY>
        </TABLE></TD></TR></TBODY></TABLE>
  
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="0">
  <tr>
    <td> <hr size="1" noshade color="#FF6600">
      <div align="center">
        <% = RsConfigObj("Copyright") %>
      </div></td>
  </tr>
</table>
<BR>
</BODY></HTML>
<%
RsConfigObj.Close
Set RsConfigObj = Nothing
Set Conn=nothing
%>