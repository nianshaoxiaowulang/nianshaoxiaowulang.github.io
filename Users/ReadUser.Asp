<% Option Explicit %>
<!--#include file="../Inc/Cls_DB.asp" -->
<!--#include file="../Inc/Const.asp" -->
<!--#include file="../Inc/Md5.asp" -->
<%
'==============================================================================
'�������ƣ���Ѷ��վ��Ϣ����ϵͳ
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
Dim DBC,conn,sConn
Set DBC = new databaseclass
Set Conn = DBC.openconnection()
Dim I,RsConfigObj
Set RsConfigObj = Conn.Execute("Select SiteName,UserConfer,Copyright,isEmail,isChange,IsShop from FS_Config")
Set DBC = Nothing
Dim RsUserObj,UserSQL
Set RsUserObj = Server.CreateObject(G_FS_RS)
UserSQL = "select * from FS_Members where MemName='"& Replace(Replace(Request("UserName"),"'",""),Chr(39),"") &"'"
RsUserObj.Open UserSQL,Conn,1,3
If RsUserObj.eof Then
	Response.Write("<script>alert(""�Ҳ����û�,�����Ѿ���ɾ����"&CopyRight&""");location=""javascript:history.back()"";</script>")  
	Response.End
Else
%>
<HTML><HEAD>
<TITLE><%=RsConfigObj("SiteName")%> >> ��Ա����</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="Css/UserCSS.css" type=text/css  rel=stylesheet>
</HEAD>
<BODY leftmargin="0" topmargin="5">
<div align="center">
  <script language="JavaScript" src="top.js" type="text/JavaScript"></script>
</div>
<TABLE width="98%" border=0 align=center cellpadding="0" cellSpacing=0>
  <TBODY>
    <TR> 
      <TD vAlign=top> <TABLE cellSpacing=0 cellPadding=0 width="100%" align=center 
                  border=0>
          <TBODY>
            <TR> 
              <TD width="100%"> <TABLE width="100%" border=0 cellpadding="0" cellspacing="0">
                  <TBODY>
                    <TR> 
                      <TD width=26><IMG 
                              src="images/Favorite.OnArrow.gif" border=0></TD>
                      <TD 
class=f4>�û�����</TD>
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
                  <div align="left"> 
                    <table width="75%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td height="3"></td>
                      </tr>
                    </table>
                    
                  <TABLE style="BORDER-COLLAPSE: collapse" borderColor=#cccccc 
            cellSpacing=0 cellPadding=0 width="100%" border=1>
                    <TBODY>
                        <TR> 
                          
                        <TD height="207" vAlign=top><table width="100%" border="0" cellpadding="8" cellspacing="1" bordercolor="#CCCCCC" bgcolor="#E6E6E6">
                            <tr bgcolor="#EEEEEE"> 
                              <td colspan="4"><font color="#FF3300"><strong>������Ϣ</strong></font></td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td><div align="right">�û���ţ�</div></td>
                              <td> <% = RsUserObj("UserNo")%> </td>
                              <td width="13%" rowspan="2"><div align="right">ͷ��</div>
                                <div align="right"></div>
                                <div align="right"></div></td>
                              <td width="39%" rowspan="2"> <%
							  If RsUserObj("HeadPic")<>"" Then
							  %> <table width="0" border="0" cellpadding="0" cellspacing="0" bgcolor="#F0F0F0">
                                  <tr> 
                                    <td bgcolor="#FFFFFF"><img src="<% = RsUserObj("HeadPic")%>" width="40" height="40" border="0"></td>
                                  </tr>
                                </table>
                                <%
							  Else
							  %> <table width="0" border="0" cellpadding="0" cellspacing="0" bgcolor="#F0F0F0">
                                  <tr> 
                                    <td bgcolor="#FFFFFF"><img src="images/noHeadPic.jpg" width="40" height="40" border="0"></td>
                                  </tr>
                                </table>
                                <%
							  End If
							  %></td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td width="17%"><div align="right">�û�����</div></td>
                              <td width="31%"><% = RsUserObj("MemName")%></td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td><div align="right">�Ա�</div></td>
                              <td> <%
								If RsUserObj("sex")=0 Then
									Response.Write("��")
								Else
									Response.Write("Ů")
								End If
								%> </td>
                              <td width="13%"><div align="right">����½ʱ�䣺</div></td>
                              <td width="39%"> <% = RsUserObj("LastLoginTime")%> </td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td><div align="right">ʡ�ݣ�</div></td>
                              <td> <% = RsUserObj("Province")%> </td>
                              <td><div align="right">���У�</div>
                                <div align="right"></div>
                                <div align="right"></div></td>
                              <td> <% = RsUserObj("City")%> </td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td><div align="right">ע��ʱ�䣺</div></td>
                              <td><% = RsUserObj("RegTime")%></td>
                              <td><div align="right">��½������</div></td>
                              <td> <% = RsUserObj("LoginNum")%> </td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td><div align="right">���ѻ��֣�</div></td>
                              <td> <% = RsUserObj("ShopPoint")%> </td>
                              <td><div align="right">�������֣�</div></td>
                              <td> <% = RsUserObj("Point")%> </td>
                            </tr>
                            <tr bgcolor="#F5F5F5"> 
                              <td height="96"> <div align="right">���ܣ�</div></td>
                              <td> <% = RsUserObj("SelfIntro")%> </td>
                              <td> <div align="right">ǩ��<strong>��</strong></div></td>
                              <td> <% = RsUserObj("UnderWrite")%> </td>
                            </tr>
                            <tr bgcolor="#EEEEEE"> 
                              <td colspan="4"> <div align="left"><font color="#FF3300"><strong>��ϵ��ʽ</strong></font></div></td>
                            </tr>
                            <%
							If RsUserObj("OpenInfTF") = 0 then
							%>
                            <tr bgcolor="#FFFFFF"> 
                              <td height="38" colspan="4"><font color="#0000FF">�û��Ѿ�����Ϊ����</font></td>
                            </tr>
                            <%
							Else
							%>
                            <tr bgcolor="#FFFFFF"> 
                              <td><div align="right">������</div></td>
                              <td colspan="3"> <% = RsUserObj("Name")%> </td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td><div align="right">�绰��</div></td>
                              <td colspan="3"> <% = RsUserObj("Telephone")%> </td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td><div align="right">��ַ��</div></td>
                              <td colspan="3"> <% = RsUserObj("address")%> </td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td><div align="right">�������룺</div></td>
                              <td colspan="3"> <% = RsUserObj("PostCode")%> </td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td><div align="right">ְҵ��</div></td>
                              <td colspan="3"> <% = RsUserObj("Vocation")%> </td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td><div align="right">ѧ����</div></td>
                              <td colspan="3"> <% = RsUserObj("EduLevel")%> </td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td><div align="right">�������ڣ�</div></td>
                              <td colspan="3"> <% = RsUserObj("birthday")%> </td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td><div align="right">��վ��ַ��</div></td>
                              <td colspan="3"> <% = RsUserObj("Homepage")%> </td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td><div align="right">QQ��</div></td>
                              <td colspan="3"><%
						if Len(RsUserObj("OICQ"))=0 Or Trim(RsUserObj("OICQ"))="" then
							Response.Write("û��")
						Else
							Dim sOICQ
						    sOICQ ="<a target=blank href=http://wpa.qq.com/msgrd?V=1&Uin="& RsUserObj("OICQ") &"&Site="& RsConfigObj("SiteName") &"&Menu=yes><img border=""0"" SRC=http://wpa.qq.com/pa?p=1:"& RsUserObj("OICQ") &":8 alt=""��������"& RsUserObj("OICQ") &"����Ϣ""></a>"
							Response.Write sOICQ
						End if
						%> </td>
                            </tr> 
                            <tr bgcolor="#FFFFFF"> 
                              <td><div align="right">MSN��</div></td>
                              <td colspan="3"> <% = RsUserObj("msn")%> </td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td><div align="right">Email:</div></td>
                              <td colspan="3"> <a href="mailto:<% = RsUserObj("email")%>"><% = RsUserObj("email")%></a></td>
                            </tr>
                            <%End If%>
                          </table> </TD>
                        </TR>
                      </TBODY>
                    </TABLE>
                    <strong></strong></div></TD>
            </TR>
          </TBODY>
        </TABLE></TD>
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
RsConfigObj.Close
Set RsConfigObj = Nothing
Set Conn=nothing
End If
%>
