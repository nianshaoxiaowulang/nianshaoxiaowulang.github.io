<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Const.asp" -->
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
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../Inc/Function.asp" -->
<%
if Request("Fid")=""  Or Not IsNumeric(Request("Fid")) Then
	Response.Write("<script>alert(""����\n����Ĳ���,����"");location.href=""javascript:history.go(-1)"";</script>")
	Response.End
end if
Fid = Replace(Replace(Request("FID"),"'",""),Chr(39),"")
Set RsConfigObj = Conn.execute("select SiteName,Copyright From FS_Config")
set Rs = server.CreateObject(G_FS_RS)
Sql = "select * from FS_Shop_Factory where IsLock=0 and id="&Fid
Rs.Open Sql,Conn,1,3
If Rs.eof then
		Response.Write("<script>alert(""����\n����Ĳ���,����,\n\n���߹���Ա�Ѿ������˴���ҵ"");location.href=""javascript:history.go(-1)"";</script>")
		Response.End
End if
%>
<html>
<title><% = Rs("CompanyName")%>__<% = RsConfigObj("SiteName") %></title>
<link href="../../CSS/FS_css.css" rel="stylesheet">
<STYLE>
.Btns {
font-family:  "����"; font-size: 30px; line-height: 35px;COLOR: #FF0000;}
.stns {
font-family:  "����"; font-size: 14px; line-height: 20px;COLOR: #003399;}
</STYLE>
<body bgcolor="#FFFFFF">
<table width="95%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#D7D7D7" class="tabbgcolor">
  <tr class="tabbgcolorliWhite"> 
    <td colspan="2" bgcolor="#FFFFFF"> <TABLE width="100%" border=0 cellpadding="5" cellspacing="0">
        <TBODY>
          <TR> 
            <TD width=26><IMG 
                              src="../../<%=UserDir%>/images/Favorite.OnArrow.gif" border=0></TD>
            <TD 
class=f4><p class="Btns"><strong> 
                <% = Rs("CompanyName")%>
                </strong></p></TD>
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
    <td height="48" colspan="2" bgcolor="#FFFFFF">
<table width="46%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><a href="ShowFactory.asp?Fid=<%=Request("FID")%>"><img src="Mall_Images/Content.gif" width="96" height="30" border="0"></a></td>
          <td><a href="ShowFactory.asp?Fid=<%=Request("FID")%>&s=Products"><img src="Mall_Images/Products.gif" width="96" height="30" border="0"></a></td>
          <td><a href="ShowFactory.asp?Fid=<%=Request("FID")%>&s=Certs"><img src="Mall_Images/Cert.gif" width="96" height="30" border="0"></a></td>
          <td><a href="ShowFactory.asp?Fid=<%=Request("FID")%>&s=Link"><img src="Mall_Images/Link.gif" width="96" height="30" border="0"></a></td>
          <td>&nbsp;</td>
        </tr>
      </table></td>
  </tr>
  <tr class="tabbgcolorliWhite"> 
    <td width="78%" height="198" colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="5">
        <tr>
          <td>
		  <%
		  If Request("s")="Products" then
		  	Call Pro()
		  ElseIf  Request("s")="Certs" then
		  	Call Cert()
		  ElseIf  Request("s")="Link" then
		  	Call Link()
		  Else
		  	Call Main()
		  End if
		  %>
		  <%
		  Sub Pro()%>
            <span class="stns">
            <% = Rs("Products")%>
            </span> 
            <%End Sub%>
		  <%
		  Sub Cert()%>
            <span class="stns">
            <% = Rs("Certs")%>
            </span> 
            <%End Sub%>
		  <%
		  Sub Link()%>
            <br>
            <br>
            <table width="73%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC" class="stns">
              <tr bgcolor="#FFFFFF"> 
                <td width="18%" bgcolor="#F5F5F5" class="stns"> <div align="right">��˾����</div></td>
                <td width="49%" class="stns"> <% = Rs("CompanyName")%> </td>
				<%if Len(Rs("Picture"))>5  then%>
                <td width="33%" rowspan="7" bgcolor="#FFFFFF" class="stns"><div align="center"><a href="<% = Rs("Picture")%>" target="_blank"><IMG SRC=<% = Rs("Picture")%> width="182" height="164" border="0"></a> 
                  </div></td> <%End iF%>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td bgcolor="#F5F5F5" class="stns"> <div align="right">��ϵ��</div></td>
                <td class="stns"> <% = Rs("LinkName")%> </td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td bgcolor="#F5F5F5" class="stns"> <div align="right">�绰</div></td>
                <td class="stns"> <% = Rs("Tel")%> </td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td bgcolor="#F5F5F5" class="stns"> <div align="right">����</div></td>
                <td class="stns"> <% = Rs("Fax")%> </td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td bgcolor="#F5F5F5" class="stns"> <div align="right">��ַ</div></td>
                <td class="stns"> <% = Rs("Address")%> </td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td bgcolor="#F5F5F5" class="stns"> <div align="right">��������</div></td>
                <td class="stns"> <% = Rs("PostCode")%> </td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td bgcolor="#F5F5F5" class="stns"> <div align="right">��վ��ַ</div></td>
                <td class="stns"> <% = Rs("HomePage")%> </td>
              </tr>
            </table> 
            <br>
            <%End Sub%>
		  <%
		  Sub Main()%>
            <span class="stns">
            <% = Rs("Content")%>
            </span> 
            <%End Sub%>
		  </td>
        </tr>
      </table> </td>
  </tr>
  <tr class="tabbgcolorliWhite"> 
    <td height="28" colspan="2" bgcolor="#EFEFEF"> 
      <div align="center"><%= RsConfigObj("Copyright")%></div></td>
  </tr>
</table>
</body></html>
<%
Set RsConfigObj = Nothing
%>