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
'�������2�ο��������뾭����Ѷ��˾������������׷����������
'==============================================================================
	Dim DBC,conn,sConn
	Set DBC = new databaseclass
	Set Conn = DBC.openconnection()
	Dim I,RsConfigObj
	Set RsConfigObj = Conn.Execute("Select Domain,SiteName,UserConfer,Copyright,isEmail,isChange,UseDatePath from FS_Config")
	Set DBC = Nothing
%>
<!--#include file="Comm/User_Purview.Asp" -->
<%
If Request.Form("Action") = "Update" then
	iF Len(Request.Form("Content"))>300 and Len(Request.Form("Content"))<=1 then
		Response.Write("<script>alert(""����\n��������Ӧ�ô���1С��300���ַ���"");location.href=""javascript:history.go(-1)"";</script>")
		Response.End
	End if
	Dim RsFobj1,RsFSQL1
	Set RsFobj1 = Server.CreateObject(G_FS_RS)
	RsFSQL1 = "Select ID,Content,Audit From FS_Review where ID="& Replace(Replace(Request.Form("ID"),"'",""),Chr(39),"")
	RsFobj1.Open RsFSQL1,Conn,1,3
	RsFobj1("Content") = Request.Form("Content")
	RsFobj1("Audit") = 0
	RsFobj1.Update
	RsFobj1.Close
	Set RsFobj1 =nothing
	Response.Write("<script>alert(""�޸ĳɹ���"&Copyright&""");location.href=""User_Comments.asp"";</script>")
	Response.End
End if
Dim RsFobj,RsFSQL
Set RsFobj = Server.CreateObject(G_FS_RS)
RsFSQL = "Select * From FS_Review where ID="& Replace(Replace(Request("ID"),"'",""),Chr(39),"")
RsFobj.Open RsFSQL,Conn,1,1
iF RsFobj("Audit")=1 Then
	Response.Write("<script>alert(""����\n��˺�����۲������޸ģ�"");location.href=""javascript:history.go(-1)"";</script>")
	Response.End
End If
iF RsFobj("UserID")<>Session("MemName") Then
	Response.Write("<script>alert(""����\n��ûȨ���޸Ĵ����ۣ�"");location.href=""javascript:history.go(-1)"";</script>")
	Response.End
End If
%>
<HTML><HEAD>
<TITLE><%=RsConfigObj("SiteName")%> >> ��Ա����</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="Css/UserCSS.css" type=text/css  rel=stylesheet>
</HEAD>
<BODY leftmargin="0" topmargin="10">
<div align="center"> </div>
<TABLE cellSpacing=2 width="98%" align=center border=0>
  <TBODY>
    <TR> 
      <form name="form2" method="post" action=""><TD vAlign=top> 
        
          <TABLE cellSpacing=0 cellPadding=5 width="98%" align=center 
                  border=0>
            <TBODY>
              <TR> 
                <TD width="100%"> <TABLE width="100%" border=0>
                    <TBODY>
                      <TR> 
                        <TD width=20><IMG src="images/Favorite.OnArrow.gif" border=0></TD>
                        <TD width="923" class=f4><p>�޸��ҷ��������</p></TD>
                      </TR>
                    </TBODY>
                  </TABLE></TD>
              </TR>
              <TR> 
                <TD width="100%"> <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                    <TBODY>
                      <TR> 
                        <TD bgColor=#ff6633 height=4><IMG height=1 src="" width=1></TD>
                      </TR>
                    </TBODY>
                  </TABLE></TD>
              </TR>
              <TR> 
                <TD width="100%" height="103" valign="top"> <div align="left"> 
                    <strong> 
                    <textarea name="Content" cols="60" rows="6" id="Content"><% = RsFobj("Content")%></textarea>
                    </strong></div></TD>
              </TR>
              <TR> 
                <TD height="31" valign="top">
<input type="submit" name="Submit" value="�޸�����">
                  <input name="Action" type="hidden" id="Action" value="Update">
                  <input name="ID" type="hidden" id="ID" value="<% = RsFobj("id")%>"></TD>
              </TR>
            </TBODY>
          </TABLE>
        </TD></form>
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
%>
<script language="JavaScript" type="text/JavaScript">
function Cim(){
	if (window.confirm('��ȷ��Ҫ����?')){
	 	return true;
	 } 
	 return false;		
}
function Cim1(){
	if (window.confirm('��ȷ��Ҫɾ����?')){
	 	return true;
	 } 
	 return false;		
}
</script>
