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
	Set Conn = DBC.openconnection()
	Dim I,RsConfigObj
	Set RsConfigObj = Conn.Execute("Select SiteName,UserConfer,Copyright,isEmail,isChange,IsShop from FS_Config")
	Set DBC = Nothing
%>
<!--#include file="Comm/User_Purview.Asp" -->
<%
	Dim RsUserObj
	Set RsUserObj = Conn.Execute("Select Point,RegTime,UserNo,UserPoint,ShopPoint From FS_Members where MemName = '"& Session("MemName")&"' and Password = '"& Session("MemPassword") &"'")
%>
<HTML><HEAD>
<TITLE><%=RsConfigObj("SiteName")%> >> ��Ա����</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="Css/UserCSS.css" type=text/css  rel=stylesheet>
</HEAD>
<BODY bgcolor="#FFFFFF" leftmargin="8" topmargin="0">
<div align="center"> </div>
<TABLE width="15%" border=0 align="center" cellpadding="0" cellSpacing=0 bgcolor="#FFFFFF">
  <TBODY>
    <TR> 
      <TD vAlign=top width=160> <a href="main_main.asp" target="main"><IMG src="images/favorite.left.help.jpg" alt="���ع���������ҳ" 
        width=190 height=27 border="0"></a>
<TABLE cellSpacing=0 cellPadding=0 width=190 border=0>
          <TBODY>
            <TR> 
              <TD> <TABLE style="BORDER-COLLAPSE: collapse" borderColor=#cccccc 
            cellSpacing=0 cellPadding=0 width="100%" border=1>
                  <TBODY>
                    <TR> 
                      <TD vAlign=top> <TABLE class=bgup cellSpacing=0 cellPadding=0 width="100%" 
                  background="" border=0>
                          <TBODY>
                            <TR> 
                              <TD width="88%" height="30" bgcolor="#FFFFFF"> <table width="75%" border="0" cellspacing="0" cellpadding="0">
                                  <tr> 
                                    <td height="1"></td>
                                  </tr>
                                </table>
                                <div align="center"> </div>
                                <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
                                  <tr> 
                                    <td height="27" bgcolor="#E8E8E8"> <div align="center">������Ϣ</div></td>
                                  </tr>
                                </table>
                                <TABLE cellSpacing=0 cellPadding=3 width="100%" 
border=0>
                                  <TBODY>
                                    <TR> 
                                      <TD height=13><A><IMG 
                              src="images/arr2.gif" width=10 height=10 id=KB1Img></A></TD>
                                      <TD><a href="All_User.Asp" target="main">ע���û�ͳ��</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD height=13 valign="top"><A><IMG 
                              src="images/arr2.gif" width=10 height=10 id=KB1Img></A></TD>
                                      <TD><a href="main_main.asp" target="main">������ȼ�</a></TD>
                                    </TR>
                                    <%
								If cint(RsConfigObj("isShop"))=1 then
								%>
                                    <TR> 
                                      <TD height=13 valign="top"><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="Mall/News.asp" target="main">��Ա����</a></TD>
                                    </TR>
                                    <%
								End if
								%>
                                    <TR> 
                                      <TD height=6 colspan="2" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                                          <tr> 
                                            <td height="27" bgcolor="#E8E8E8"> 
                                              <div align="center">�ʺ���Ϣ</div></td>
                                          </tr>
                                        </table></TD>
                                    </TR>
                                    <TR> 
                                      <TD width=14 height=13 valign="top"><A><IMG 
                              src="images/arr2.gif" width=10 height=10 id=KB1Img></A></TD>
                                      <TD> <a href="User_Modify_account.asp" target="main">�ʺ���Ϣ</a>/<a href="User_Modify_contact.asp" target="main">��ϵ��ʽ</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD height=13><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="User_Modify_Pass.asp" target="main">�޸�������ʾ��</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD width=14 height=13><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="User_Modify_other.asp" target="main">������ϵ��ʽ</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD height=6 colspan="2" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                                          <tr> 
                                            <td height="27" bgcolor="#E8E8E8"> 
                                              <div align="center">��Ϣ����</div></td>
                                          </tr>
                                        </table></TD>
                                    </TR>
                                    <TR> 
                                      <TD height=13><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="User_Comments.asp" target="main">�ҷ��������</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD height=13><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="User_Favorite.asp" target="main">���ղص���Ϣ</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD height=13><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="User_contribution.asp" target="main">�������</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD height=13><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="Add_Contribution.asp" target="main">��Ӹ��</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD height=13><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="GBook/All_GBook.asp" target="main">���Թ���</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD height=13 colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                                          <tr> 
                                            <td height="27" bgcolor="#E8E8E8"> 
                                              <div align="center">����Ϣ</div></td>
                                          </tr>
                                        </table></TD>
                                    </TR>
                                    <TR> 
                                      <TD height=13><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="User_WriteMessage.asp" target="main">׫д��Ϣ</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD height=13><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="User_Message.asp?action=Inbox" target="main">�ռ���</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD height=13><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="User_Message.asp?action=Outbox" target="main">������</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD height=13><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="User_Message.asp?action=Recycle" target="main">�ϼ���</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD height=13><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="User_AddressList.asp" target="main">��ַ��</a></TD>
                                    </TR>
                                  </TBODY>
                                </TABLE>
                                <%
								If cint(RsConfigObj("isShop"))=1 then
								%> <table width="98%" border="0" cellspacing="0" cellpadding="3">
                                  <tr> 
                                    <td height="27" bgcolor="#E8E8E8"> <div align="center">�̳ǹ���</div></td>
                                  </tr>
                                </table>
                                <TABLE cellSpacing=0 cellPadding=3 width="100%" 
border=0>
                                  <TBODY>
                                    <TR> 
                                      <TD width=14 height=13><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD width="154"> <a href="Mall/BuyOrder.asp" target="main">��������</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD width=14 height=13><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="Mall/Integral.asp" target="main">�ҵĻ���/���</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD width=14 height=13><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="Mall/Favorite.asp" target="main">�ղؼ�</a></TD>
                                    </TR>
                                    <TR style="display:"> 
                                      <TD height=13><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="Mall/BuyProductPack.asp" target="main">���ﳵ</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD height=13><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="Mall/Exchange.asp" target="main">���ֻ����</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD height="7"><IMG height=5 src="images/SelfService.aspx" 
                              width=1><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="Mall/RegCompany.asp" target="main">ע���ҵ���ҵ</a></TD>
                                    </TR>
                                    <TR> 
                                      <TD height="7"><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><a href="Mall/RegCompanyManage.asp" target="main">�޸��ҵ���ҵ</a></TD>
                                    </TR>
                                    <TR>
                                      <TD height="7"><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><A href="mall/Pmf.asp" target="main">������֪</A></TD>
                                    </TR>
                                  </TBODY>
                                </TABLE>
                                <%
								Else
								%> <table width="98%" border="0" align="center" cellpadding="3" cellspacing="0">
                                  <tr> 
                                    <td height="27" bgcolor="#E8E8E8"> <div align="center">�̳ǹ���</div></td>
                                  </tr>
                                </table>
                                <TABLE width="100%" height="27" 
border=0 cellPadding=3 cellSpacing=0>
                                  <TBODY>
                                    <TR> 
                                      <TD height="21"><font color="#FF0000">δ��ͨ</font></TD>
                                    </TR>
                                  </TBODY>
                                </TABLE>
                                <%
								End if
								%> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                  <tr> 
                                    <td height="27" bgcolor="#E8E8E8"> <div align="center"><a href="main_main.asp" target="main"><font color="#FF0000">��Ա������ҳ</font></a>��<a href="Comm/LetOut.asp" target="_top"><font color="#990000">��ȫ�˳�</font></a></div></td>
                                  </tr>
                                </table></TD>
                            </TR>
                          </TBODY>
                        </TABLE></TD>
                    </TR>
                  </TBODY>
                </TABLE></TD>
            </TR>
          </TBODY>
        </TABLE></TD>
    </TR>
  </TBODY>
</TABLE>
  
<table width="100%" border="0" cellspacing="3" cellpadding="0">
  <tr> 
    <td height="5">��Ѷ�ٷ�վ��<a href="http://www.Foosun.Cn" target="_blank">Foosun.Cn</a> 
    </td>
  </tr>
  <tr> 
    <td height="5">��Ѷ����վ��<a href="http://Help.Foosun.Net" target="_blank">Help.Foosun.Net</a></td>
  </tr>
  <tr> 
    <td height="5">��Ѷ����վ��<a href="http://BBS.Foosun.Net" target="_blank">BBS.Foosun.Net</a></td>
  </tr>
</table>
</BODY></HTML>
<%
Sub SendEmail()

End Sub
Sub EmailInfo()
	Response.Write("һ�����Ѿ����͵���ע��ĵ����ʼ�<font color=red>"& Session("email") &"</font>����ע����գ�")
End Sub
RsConfigObj.Close
Set RsConfigObj = Nothing
RsUserObj.close
Set RsUserObj=nothing
Set Conn=nothing
%>