<% Option Explicit %>
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
'�������2�ο��������뾭����Ѷ��˾������������׷����������
'==============================================================================
Dim DBC,conn,sConn
Set DBC = new databaseclass
Set Conn = DBC.openconnection()
Set DBC = Nothing
'��鰲װ�ļ��Ƿ���ڣ�������ڣ�ת�򵽰�װĿ¼
Dim FileObj  
Set FileObj=Server.CreateObject(G_FS_FSO)
If FileObj.FileExists(Server.MapPath("../../Install/install.asp")) = True then
	Response.Write("<script>alert(""ϵͳ���ڰ�װ�ļ����밲װ���½"&Copyright&""");location.href=""../../Install/install.asp"";</script>")
	Response.End
End if
Dim Action
Action = "CheckLogin.ASP?UrlAddress=" & Request("UrlAddress")
Function GetCode()
	Dim TestObj
	On Error Resume Next
	Set TestObj = Server.CreateObject("Adodb.Stream")
	Set TestObj = Nothing
	If Err Then
		Dim TempNum
		Randomize timer
		TempNum = cint(8999*Rnd+1000)
		Session("GetCode") = TempNum
		GetCode = Session("GetCode")		
	Else
		GetCode = "<img src=""GetCode.asp"" onclick='this.src=this.src;' style='cursor:pointer'>"		
	End If
End Function
Function GetSiteName
	Dim RsConfigLoginobj
	On error resume next
	Set RsConfigLoginobj=Conn.execute("Select SiteName from FS_Config")
	If not RsConfigLoginobj.eof then
		GetSiteName=RsConfigLoginobj("SiteName")
	Else
		GetSiteName="��Ѷ"
	End If
	If err.number<>0 then GetSiteName="��Ѷ"
	Set RsConfigLoginobj = Nothing
End Function
%>
<HTML><HEAD>
<TITLE><% = GetSiteName %>___��վ���ݹ���ϵͳ___��̨��¼</TITLE>
<META http-equiv=Content-Type content=text/html;charset=gb2312>
<STYLE>
.tableborder {
	BORDER-RIGHT: #737373 1px solid; BORDER-TOP: #bbbbbb 1px solid; BORDER-LEFT: #bbbbbb 1px solid; BORDER-BOTTOM: #737373 1px solid; BACKGROUND-COLOR: #d8dbd7
}
.setupheader {
	FONT-WEIGHT: bold; FONT-SIZE: 14px; COLOR: #ffffff; BACKGROUND-COLOR: #454545
}
.button {
	FONT-SIZE: 12px; CURSOR: pointer; COLOR: #000000; FONT-FAMILY: Tahoma, Verdana, Arial; HEIGHT: 22px
}
.topheader {
	PADDING-RIGHT: 3px; PADDING-LEFT: 3px; FONT-WEIGHT: bold; PADDING-BOTTOM: 3px; COLOR: #ffffff; PADDING-TOP: 3px; BACKGROUND-COLOR: #336699
}
.header_box {
	PADDING-RIGHT: 1px; PADDING-LEFT: 1px; PADDING-BOTTOM: -1px; VERTICAL-ALIGN: middle; PADDING-TOP: 1px; HEIGHT: 1px; BACKGROUND-COLOR: #ffffff;
}
.header_box1 {
	PADDING-RIGHT: 1px; PADDING-LEFT: 1px; PADDING-BOTTOM: 1px; VERTICAL-ALIGN: middle; PADDING-TOP: 1px; HEIGHT: 1px; BACKGROUND-COLOR: #ffffff;
}
.install_box {
	PADDING-RIGHT: 1px; PADDING-LEFT: 1px; PADDING-BOTTOM: 1px; PADDING-TOP: 1px; BACKGROUND-COLOR: #d4d0c8
}
.firsthr {
	BACKGROUND-COLOR: #808080
}
.secondhr {
	BACKGROUND-COLOR: #ffffff
}
td {
	font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
	font-size:12px;
}
.STYLE1 {color: #FFFEFF}
</STYLE>
<META content="MSHTML 6.00.3790.2491" name=GENERATOR>
<meta name="keywords" content="http://www.skyim.com/">
</HEAD>
<BODY topMargin=30>
<TABLE class=tableborder cellSpacing=1 cellPadding=0 width=496 align=center border=0>
		  <form action="checklogin.asp" method="post">
  <TBODY>
  <TR>
    <TD>
      <DIV class=topheader>&nbsp;&nbsp;<a href="http://www.skyim.com/" target="_blank" class="STYLE1">www.skyim.com</a>��̨��¼ ��������������������������<font color="#CCCCCC">�汾�ţ�V1.0</font></DIV>
      <DIV class=header_box><a href="http://www.skyim.com/" target="_blank"><img src="images/login.jpg" border="0"></a></DIV>
      <DIV class=install_box>
		    <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td colspan="3">&nbsp;</td>
              </tr>
              <tr> 
                <td width="48%" align="right">�û���: 
                  <input name="name" type="text" style="width:92px" value="<%=Request.Cookies("FoosunCookie")("AdminName")%>"></td>
                <td width="3%" align="left">&nbsp;</td>
                <td width="48%" align="left"> ��ס�û��� 
                  <input name="AutoGet" type="checkbox" id="AutoGet" value="1" <% If Request.Cookies("FoosunCookie")("AdminName")<>"" Then Response.Write "checked" End If%>></td>
              </tr>
              <tr> 
                <td align="right"> �ܡ���: 
                  <input name="password" type="password" style="width:92px;FONT-SIZE:12px;"></td>
                <td align="left">&nbsp;</td>
                <td align="left"> �顡֤����: 
                  <input name="VerifyCode" type="text" size="4"> 
                  <% = GetCode() %>
                </td>
              </tr>
              <tr> 
                <td>&nbsp;</td>
              </tr>
            </table>
      <TABLE cellSpacing=0 cellPadding=2 width="100%" border=0>
        <TBODY>
        <TR>
                  <TD style="FONT-SIZE:11px;COLOR:#666666" noWrap width="4%">http://www.skyim.com&nbsp;</TD>
          <TD>
            <DIV class=firsthr style="HEIGHT: 1px"><IMG height=1 alt="" src="" 
            width=1></DIV>
            <DIV class=secondhr style="HEIGHT: 1px"><IMG height=1 alt="" src="" 
            width=1></DIV></TD></TR></TBODY></TABLE>
    <DIV align=right>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td align="right">
		  	<input class=button  type=submit value='  ��¼  '>&nbsp;&nbsp;
            <input class=button onClick="javascript:window.opener=null;window.close();" type=button value='  �ر�   '>
            &nbsp;&nbsp;</td>
        </tr>
        <tr>
          <td align="right">&nbsp;</td>
        </tr>
      </table>
    </DIV></DIV></TD></TR>
</TBODY>
</form>
</TABLE>
</BODY></HTML>
<script language="JavaScript">
CheckBrowerVersion();
var ErrInfo='<% = Request("ErrInfo")%>';
function CheckBrowerVersion()
{
	var MajorVer=navigator.appVersion.match(/MSIE (.)/)[1];
	var MinorVer=navigator.appVersion.match(/MSIE .\.(.)/)[1];
	var IE6OrMore=MajorVer>= 5.5||(MajorVer>=5.5&&MinorVer>=5.5);
	if (!IE6OrMore)
	{
		alert('IE������汾̫�ͣ�ϵͳ�������������С����ȷ�������㵽���ص�ַ��');
		location.href="http://nj.onlinedown.net/soft/17441.htm"
		//document.all.BtnSubmit.disabled=true;
	}
}
if (ErrInfo!='') alert(ErrInfo);
</script>

