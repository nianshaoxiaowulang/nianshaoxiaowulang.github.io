<% Option Explicit %>
<!--#include file="../Inc/Cls_DB.asp" -->
<!--#include file="../Inc/Const.asp" -->
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
	Set RsConfigObj = Conn.Execute("Select SiteName,UserConfer,Copyright from FS_Config")
%>
<HTML><HEAD><TITLE><%=RsConfigObj("SiteName")%> >> ��д�ʺ���Ϣ</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="Css/UserCSS.css" type=text/css  rel=stylesheet></HEAD>
<BODY leftmargin="0" topmargin="10">
<div align="center">
  <script language="JavaScript" src="top.js" type="text/JavaScript"></script>
  <script language="javascript" src="Comm/MyScript.js"></script>
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
                                      <TD><font color="#FF0000">��д�ʺ���Ϣ</font></TD>
                                    </TR>
                                    <TR> 
                                      <TD width=14 height=24><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD>��д��ϵ����</TD>
                                    </TR>
                                    <TR> 
                                      <TD width=14 height=24><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD>ע��ɹ�</TD>
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
class=f4><span class="f4">��д�ʺ���Ϣ</span></TD>
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
              <form method=POST action="sRegister_1.asp" name=UserForm onSubmit="return checkdata()">
                <TD width="100%" height="74"> <div align="left"> <br>
                    <table width="98%" border="0" cellspacing="0" cellpadding="5">
                      <tr> 
                        <td width="19%"><div align="right">�û�����</div></td>
                        <td width="29%"><input name="Username" type="text" id="Username">
                          <font color="#FF0000"> * <a href="javascript:CheckName('Comm/CheckName.asp')"><font color="#999900"><br>
                          ����Ƿ�ռ��</font></a> </font></td>
                        <td width="52%">��¼ʱʹ�õĴ��ţ����ס�ÿͻ���������ʹ��email��ַ���Ա���䣡 </td>
                      </tr>
                      <tr> 
                        <td><div align="right">���룺</div></td>
                        <td><input name="sPassword" type="password" id="sPassword"> 
                          <font color="#FF0000">*</font></td>
                        <td>Ϊ�����ĸ������ϵİ�ȫ���������ʹ��6λ�������룡 </td>
                      </tr>
                      <tr> 
                        <td><div align="right">��֤���룺</div></td>
                        <td><input name="Confimpass" type="password" id="Confimpass"> 
                          <font color="#FF0000">*</font></td>
                        <td>������������������һ�£� </td>
                      </tr>
                      <tr> 
                        <td><div align="right">�����ʼ���</div></td>
                        <td><input name="email" type="text" id="email"> <font color="#FF0000">*<br>
                          <a href="javascript:CheckEmail('Comm/Checkemail.asp')"><font color="#999900">����Ƿ�ռ��</font></a></font></td>
                        <td>���������ʹ�õĵ����ʼ���ַ��ÿ��Emailֻ��ע��һλ�û����� </td>
                      </tr>
                      <tr> 
                        <td><div align="right">������ʾ���⣺</div></td>
                        <td><input name="PassQuestion" type="text" id="PassQuestion">
                          <font color="#FF0000">*</font></td>
                        <td>������ʾ���⣬������ʾ���һ�����ĸ���</td>
                      </tr>
                      <tr> 
                        <td><div align="right">������ʾ�𰸣�</div></td>
                        <td><input name="PassAnswer" type="text" id="PassAnswer">
                          <font color="#FF0000">*</font></td>
                        <td>������ʾ�𰸣�������ʾ���һ������Ψһ����֮һ</td>
                      </tr>
                      <tr> 
                        <td>&nbsp;</td>
                        <td colspan="2"><input  type=submit name="Submit3" value="��һ��" style="cursor:hand;"></td>
                      </tr>
                      <tr> 
                        <td>&nbsp;</td>
                        <td colspan="2">ע�⣺<br>
                          1����*����Ŀ������д������ע�᲻�ܼ����� <br>
                          2���Ƽ���ʹ������2G�����������@126.com����� <a href="http://reg.126.com/reg1.jsp" target="_blank"><font color="#FF0000">����ע��</font></a> 
                        </td>
                      </tr>
                    </table>
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
<script language="JavaScript" type="text/JavaScript">
function CheckName(gotoURL) {
   var ssn=UserForm.Username.value.toLowerCase();
	   var open_url = gotoURL + "?Username=" + ssn;
	   window.open(open_url,'','status=0,directories=0,resizable=0,toolbar=0,location=0,scrollbars=0,width=150,height=80');
}
function CheckEmail(gotoURL) {
   var ssn1=UserForm.email.value.toLowerCase();
	   var open_url = gotoURL + "?email=" + ssn1;
	   window.open(open_url,'','status=0,directories=0,resizable=0,toolbar=0,location=0,scrollbars=0,width=150,height=80');
}
</script>
