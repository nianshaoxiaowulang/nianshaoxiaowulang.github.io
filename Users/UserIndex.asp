<!--#include file="../Inc/Cls_DB.asp" -->
<!--#include file="../Inc/Const.asp" -->
<!--#include file="../Inc/Md5.asp" -->
<!--#include file="../Inc/NoSqlHack.asp" -->
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
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
if Request.Form("action")="Check" then
	MemName = Replace(Trim(Request.Form("MemName")),"'","''")
	Password = md5(Request.Form("MemPass"),16)

if MemName = "" or  Password = "" then 
	Response.write"<script>alert(""�û��������벻��Ϊ��"");location.href=""javascript:history.back()"";</script>"
	Response.end
end if
Set RsLoginObj = Server.CreateObject(G_FS_RS)
SqlLogin = "Select * from FS_members where MemName='"&MemName&"' and  password='"&Password&"'"
RsLoginObj.Open SqlLogin,Conn,1,1
if Not RsLoginObj.EOF then 
   if RsloginObj("Lock")=true then
	   Response.write"<script>alert(""���Ѿ�������������ϵ����Ա"");location.href=""javascript:history.back()"";</script>"
	   Response.end
   end if
   Response.Cookies("Foosun")("MemName") = MemName
   Response.Cookies("Foosun")("MemPassword") = Password
   Response.Cookies("Foosun")("MemID") = RsLoginObj("ID")
   Response.Cookies("Foosun")("GroupID") = RsLoginObj("GroupID")
   Session("MemName")=MemName
   Session("MemPassword")=Password
   Session("MemID")=RsLoginObj("ID")
   dim LoginTime
   LoginTime = Now()
   conn.execute("Update FS_members set LoginNum=LoginNum+1,Point=Point+1,LastLoginIP='"&Request.ServerVariables("Remote_ADDR")&"',LastLoginTime='"&LoginTime&"' where MemName='"&MemName&"'")'�û���½һ�Σ�����+1��
   Response.Redirect("UserIndex.asp") 
   Response.End
else
   Response.write"<script>alert(""�Ƿ���½�������û������������ȷ��"");location.href=""javascript:history.back()"";</script>"
   Response.end
end if
set Conn = Nothing
Set RsLoginObj = Nothing

end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��Ա����</title>
<style type="text/css">
<!--
 BODY   {border: 0; margin: 0; cursor: default; font-family:����; font-size:9pt;}
 BUTTON {width:5em}
 TABLE  {font-family:����; font-size:9pt}
 P      {text-align:center}
-->
</style>
</head>
<body leftmargin="0" topmargin="0">
<%
MemName = Session("MemName")
PassWord = Session("MemPassword")
MemID = Session("MemID")
set RsMemObj = Server.CreateObject (G_FS_RS)
RsMemObj.Source="select * from FS_Members where MemName='"& MemName &"' and password='"&PassWord&"'"
RsMemObj.Open RsMemObj.Source,Conn,1,1
if not RsMemObj.EOF then
%>
<table width="226" border="0" cellpadding="2" cellspacing="0">
  <tr> 
    <td colspan="4" class="tabbgcolorlileft"><span class="Nred9pt"><font color="#FF0000"><%=MemName%></font></span><font color="#FF0000">��</font>��ӭ����<a href="Main.asp" target="_top">�������</a> 
      <a href="Comm/LetOut.asp" target="_top">�˳�</a> </td>
  </tr>
  <tr> 
    <td width="67"> <div align="right">һ����֣�</div></td>
    <td width="36"><%=RsMemObj("Point")%></td>
    <td width="60">��½������</td>
    <td width="47"><%=RsMemObj("LoginNum")%></td>
  </tr>
  <tr> 
    <td width="67"> <div align="right">ע��ʱ�䣺</div></td>
    <td colspan="3"><%=RsMemObj("RegTime")%></td>
  </tr>
  <tr> 
    <td width="67"> <div align="right">��½ʱ�䣺</div></td>
    <td colspan="3"><%=RsMemObj("LastLoginTime")%></td>
  </tr>
  <tr> 
    <td width="67"> <div align="right">��½�ɣУ�</div></td>
    <td colspan="3"><%=RsMemObj("LastLoginIP")%></td>
  </tr>
  <tr> 
    <td colspan="4"><div align="center"><a href="User_Modify_Pass.asp" target="_top">�޸�����</a>����<a href="User_Modify_contact.asp" target="_top">�޸�����</a>��</div></td>
  </tr>
</table>
<%
else
%>
<table width="100%" border="0" cellspacing="5" cellpadding="0">
  <form action="" method="post" name="LoginForm">
    <tr> 
      <td> 
        <div align="center">�û��� 
          <input name="MemName" type="text" id="MemName" size="15">
      </div></td>
    </tr>
    <tr> 
      <td> 
        <div align="center">���룺 
          <input name="MemPass" type="password" id="MemPass" size="15">
      </div></td>
    </tr>
    <tr> 
      <td><div align="center">
          <input type="submit" name="Submit" value="��¼">
          <input type="reset" name="Submit2" value="����">
          <input name="action" type="hidden" id="action" value="Check">&nbsp;&nbsp;
          <a href="Register.asp" target="_top"><font color="#FF0000">ע��</font></a> 
          <a href="User_GetPassword.asp" target="_top">��������</a></div></td>
    </tr>
  </form>
</table>
  <%
end if
%>
</body>
</html>
