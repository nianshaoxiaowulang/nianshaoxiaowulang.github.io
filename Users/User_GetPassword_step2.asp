<% Option Explicit %>
<!--#include file="../Inc/Cls_DB.asp" -->
<!--#include file="../Inc/Const.asp" -->
<!--#include file="../Inc/Md5.asp" -->
<!--#include file="../Inc/Function.asp" -->
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
	Set RsConfigObj = Conn.Execute("Select SiteName,Copyright from FS_Config")
If request.form("action")="modify" then
	If  Replace(Trim(Request.Form("password")),"'","''") =""  then 
			Response.Write("<script>alert(""���벻��Ϊ�գ�"");location=""Javascript:history.go(-1)"";</script>")
			Response.End
	End if
	If  Replace(Trim(Request.Form("password")),"'","''") <> Replace(Trim(Request.Form("password1")),"'","''")  then
			Response.Write("<script>alert(""2�����벻��ͬ�����������룡"");location=""Javascript:history.go(-1)"";</script>")
			Response.End
	End if
	Set RsMemObj = Server.CreateObject (G_FS_RS)
	RsMemObj.Source="select * from FS_Members where ID="& CLNG(Replace(Request.Form("MemID"),Chr(39),""))
	RsMemObj.Open RsMemObj.Source,Conn,1,3
	RsMemObj("Password") = md5(Replace(Trim(Request.Form("password")),"'","''"),16)
	RsMemObj.update
	Response.Write("<script>alert(""�޸�����ɹ���"");location=""Login.asp"";</script>")  
	Response.End
End if
%>
<HTML><HEAD><TITLE><%=RsConfigObj("SiteName")%> >> ��Ա��½</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<META content="MSHTML 6.00.2600.0" name=GENERATOR>
<META content=JavaScript name=vs_defaultClientScript>
<LINK href="Css/UserCSS.css" type=text/css rel=stylesheet>
</HEAD>
<script language="javascript" src="Comm/MyScript.js"></script>
<BODY leftMargin=0 topMargin=10 MS_POSITIONING="GridLayout">
<div align="center">
  <script language="JavaScript" src="top.js" type="text/JavaScript"></script>
</div>
<TABLE cellSpacing=2 width="98%" align=center border=0>
  <TBODY>
    <TR> 
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
class=f4>��Ա��½</TD>
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
              <TD width="100%" height="37"> <div align="left"> <br>
<%
Dim RsMemObj
set RsMemObj = Server.CreateObject (G_FS_RS)
RsMemObj.Source="select * from FS_Members where MemName='"& Replace(Trim(Request.Form("MemName")),"'","''") &"' and PassAnswer='"& MD5(Replace(Trim(Request.Form("PassAnswer")),"'","''"),16) &"'and email='"& Replace(Trim(Request.Form("email")),"'","''") &"'"
RsMemObj.Open RsMemObj.Source,Conn,1,1
If Not RsMemObj.eof then
	Call main()
Else
	Call Nonerecode()
End if
Sub main()%>
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td>&nbsp;</td>
                    </tr>
                  </table>
                  <form name="form1" method="post" action="">
                    <table width="44%" border="0" align="center" cellpadding="5" cellspacing="1" bordercolor="e6e6e6" bgcolor="#B6B6B6">
                      <tr bgcolor="#EEEEEE"> 
                        <td width="23%"> 
                          <div align="right">�޸����룺</div></td>
                        <td width="77%"><font color="#FF0000"> 
                          <input name="action" type="hidden" id="action" value="modify">
                          <input name="MemID" type="hidden" value="<%=RsMemObj("ID")%>">
                          <input name="Password" type="password" id="Password">
                          </font><a href="UserLogin.asp"><font color="#0000FF"> 
                          </font></a></td>
                      </tr>
                      <tr bgcolor="#EEEEEE"> 
                        <td> 
                          <div align="right">ȷ�����룺</div></td>
                        <td><font color="#FF0000"> 
                          <input name="Password1" type="password" id="Password1">
                          </font></td>
                      </tr>
                      <tr bgcolor="#EEEEEE"> 
                        <td>&nbsp;</td>
                        <td><a href="UserLogin.asp"><font color="#0000FF"> 
                          <input type="submit" name="Submit" value="�޸�">
                          </font></a></td>
                      </tr>
                    </table>
                  </form>
<%
End sub
Sub Nonerecode()
%>
                  <table width="80%" border="0" align="center" cellpadding="5" cellspacing="1" bordercolor="e6e6e6" bgcolor="e6e6e6">
                    <tr>
                      <td bgcolor="#FFFFFF"> 
                        <%
						Response.Write("���������ϲ���ȷ�����������룡")
						%>
                      </td>
                    </tr>
                  </table>
<%
End sub
%>
                </div></TD>
            </TR>
          </TBODY>
        </TABLE></TD>
    </TR>
  </TBODY>
</TABLE>
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="0">
  <tr> 
    <td> <hr size="1" noshade color="#FF6600"> <div align="center"> 
        <% = RsConfigObj("Copyright") %>
      </div></td>
  </tr>
</table>
</BODY></HTML>
