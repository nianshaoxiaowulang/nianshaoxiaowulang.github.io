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
Dim DBC,conn,sConn
Set DBC = new databaseclass
Set Conn = DBC.openconnection()
Dim I,RsConfigObj
Set RsConfigObj = Conn.Execute("Select SiteName,UserConfer,Copyright,isEmail,isChange,IsShop from FS_Config")
Set DBC = Nothing
%>
<!--#include file="Comm/User_Purview.Asp" -->
<HTML><HEAD>
<TITLE><%=RsConfigObj("SiteName")%> >> ��Ա����</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="Css/UserCSS.css" type=text/css  rel=stylesheet>
</HEAD>
<BODY leftmargin="0" topmargin="5">
<div align="center"> </div>
<TABLE cellSpacing=2 width="98%" align=center border=0>
  <TBODY>
    <TR> 
      <TD vAlign=top> <TABLE cellSpacing=0 cellPadding=0 width="98%" align=center 
                  border=0>
          <TBODY>
            <TR> 
              <TD width="100%"> <TABLE width="100%" border=0 cellpadding="0" cellspacing="0">
                  <TBODY>
                    <TR> 
                      <TD width=26><IMG 
                              src="images/Favorite.OnArrow.gif" border=0></TD>
                      <TD 
class=f4>ѡ��ͷ��</TD>
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
              <form method=POST action="" name="UserForm1">
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
                          <TD vAlign=top> <TABLE class=bgup cellSpacing=0 cellPadding=5 width="100%" 
                  background="" border=0>
                              <TBODY>
                                <TR> 
                                  <TD width="95%" height="68">
									<table width="100%" border="0" cellspacing="2" cellpadding="6">
                                      <tr>
									  <%
									  Dim m,j
									  j=1
									  For m = 1 to 99
									  %>
                                        <td align="center"><IMG SRC="Images/HeadPicture/<%=m%>.gif" onclick="insertface('Images/HeadPicture/<%=m%>.gif',this.width,this.height)" style="CURSOR: hand" ></a><br><%=m%>.gif</td></td>
									  <%
									  if j mod 6 = 0 then
										  Response.Write("</tr>")
									  End If
									  j=j+1
									  Next
									  %>
                                    </table>
									</TD>
                                </TR>
                              </TBODY>
                            </TABLE></TD>
                        </TR>
                      </TBODY>
                    </TABLE>
                    <strong></strong></div></TD>
              </form>
            </TR>
          </TBODY>
        </TABLE><br><div align="center"><BUTTON onclick="window.close();">�ر�</BUTTON></div></TD>
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
<script language="javascript">
function insertface(userface,width,height){
	self.opener.UserForm1.HeadPic.value=userface;
	self.close()
}
</script>
</BODY></HTML>
<%
RsConfigObj.Close
Set RsConfigObj = Nothing
Set Conn=nothing
%>

