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
              <TD width="100%" height="37"> <div align="left"> 
                  <form name=LoginForm method="post" action="User_GetPassword_step2.asp" onSubmit="return CheckLogindata()">
                    <%
Dim RsMemObj
Set RsMemObj = Server.CreateObject (G_FS_RS)
RsMemObj.Source="select * from FS_Members where MemName='"& Replace(Replace(Trim(Request.Form("MemName")),"'","''"),Chr(39),"") &"'"
RsMemObj.Open RsMemObj.Source,Conn,1,1
If Not RsMemObj.eof then
	Call main()
Else
	Call Nonerecode()
End if
Sub main()
%>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td>&nbsp;</td>
                      </tr>
                    </table>
                    <div align="center"> 
                      <table width="44%" border="0" align="center" cellpadding="5" cellspacing="1" bordercolor="e6e6e6" bgcolor="#B4B4B4">
                        <tr bgcolor="#F0F0F0"> 
                          <td width="39%"> 
                            <div align="right">����������⣺</div></td>
                          <td width="61%"> 
                            <input name="MemName" type="hidden" id="MemName" value="<%=RsMemObj("MemName")%>"> 
                            <%=RsMemObj("PassQuestion")%></td>
                        </tr>
                        <tr bgcolor="#F0F0F0"> 
                          <td> 
                            <div align="right">�������������𰸣�</div></td>
                          <td> 
                            <input name="PassAnswer" type="password" id="PassAnswer"></td>
                        </tr>
                        <tr bgcolor="#F0F0F0"> 
                          <td> 
                            <div align="right">��ĵ����ʼ���</div></td>
                          <td> 
                            <input name="email" type="text" id="PassAnswer3"></td>
                        </tr>
                        <tr bgcolor="#F0F0F0"> 
                          <td colspan="2"> 
                            <div align="center"> 
                              <input name="Submit" type="submit" class="Anbut1" value="��һ��">
                            </div></td>
                        </tr>
                      </table>
                    </div>
<%
End sub
Sub Nonerecode()
%>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td>&nbsp;</td>
                      </tr>
                    </table>
                    <table width="80%" border="0" align="center" cellpadding="5" cellspacing="1" bordercolor="e6e6e6" bgcolor="e6e6e6">
                      <tr>
                        <td bgcolor="#FFFFFF"> 
<%
Response.Write("δ��ѯ�����û�")
%>
                        </td>
                      </tr>
                    </table>
<%
End sub
%>
                    <div align="center"></div>
                  </form>
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
