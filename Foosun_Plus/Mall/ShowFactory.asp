<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Const.asp" -->
<% 
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
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../Inc/Function.asp" -->
<%
if Request("Fid")=""  Or Not IsNumeric(Request("Fid")) Then
	Response.Write("<script>alert(""错误：\n错误的参数,请检查"");location.href=""javascript:history.go(-1)"";</script>")
	Response.End
end if
Fid = Replace(Replace(Request("FID"),"'",""),Chr(39),"")
Set RsConfigObj = Conn.execute("select SiteName,Copyright From FS_Config")
set Rs = server.CreateObject(G_FS_RS)
Sql = "select * from FS_Shop_Factory where IsLock=0 and id="&Fid
Rs.Open Sql,Conn,1,3
If Rs.eof then
		Response.Write("<script>alert(""错误：\n错误的参数,请检查,\n\n或者管理员已经锁定了此企业"");location.href=""javascript:history.go(-1)"";</script>")
		Response.End
End if
%>
<html>
<title><% = Rs("CompanyName")%>__<% = RsConfigObj("SiteName") %></title>
<link href="../../CSS/FS_css.css" rel="stylesheet">
<STYLE>
.Btns {
font-family:  "黑体"; font-size: 30px; line-height: 35px;COLOR: #FF0000;}
.stns {
font-family:  "宋体"; font-size: 14px; line-height: 20px;COLOR: #003399;}
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
                <td width="18%" bgcolor="#F5F5F5" class="stns"> <div align="right">公司名称</div></td>
                <td width="49%" class="stns"> <% = Rs("CompanyName")%> </td>
				<%if Len(Rs("Picture"))>5  then%>
                <td width="33%" rowspan="7" bgcolor="#FFFFFF" class="stns"><div align="center"><a href="<% = Rs("Picture")%>" target="_blank"><IMG SRC=<% = Rs("Picture")%> width="182" height="164" border="0"></a> 
                  </div></td> <%End iF%>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td bgcolor="#F5F5F5" class="stns"> <div align="right">联系人</div></td>
                <td class="stns"> <% = Rs("LinkName")%> </td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td bgcolor="#F5F5F5" class="stns"> <div align="right">电话</div></td>
                <td class="stns"> <% = Rs("Tel")%> </td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td bgcolor="#F5F5F5" class="stns"> <div align="right">传真</div></td>
                <td class="stns"> <% = Rs("Fax")%> </td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td bgcolor="#F5F5F5" class="stns"> <div align="right">地址</div></td>
                <td class="stns"> <% = Rs("Address")%> </td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td bgcolor="#F5F5F5" class="stns"> <div align="right">邮政编码</div></td>
                <td class="stns"> <% = Rs("PostCode")%> </td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td bgcolor="#F5F5F5" class="stns"> <div align="right">网站地址</div></td>
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