<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Md5.asp" -->
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
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P040204") then Call ReturnError()
Dim AdminID,Result
Result = Request("Result")
AdminID = Replace(Replace(Request("AdminID"),"'",""),"""","")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�޸Ĺ���Ա����</title>
</head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet">
<body>
<div align="center">
  <table width="95%" border="0" cellspacing="0" cellpadding="0">
  <form action="" method="post" name="PassWordForm">
    <tr>
      <td height="15" colspan="2">&nbsp;</td>
      </tr>
    <tr> 
      <td width="23%" height="30"> 
        <div align="left">&nbsp;&nbsp;&nbsp;�� �� ��</div></td>
      <td width="77%" height="30"> 
        <div align="left"> 
          <input name="PassWord" type="password" id="PassWord" style="width:100%;">
        </div></td>
    </tr>
    <tr> 
      <td height="30"><div align="left">&nbsp;&nbsp;&nbsp;ȷ������</div></td>
      <td height="30"><div align="left"><input name="AffirmPassWord" type="password" id="AffirmPassWord" style="width:100%;">
        </div></td>
    </tr>
    <tr> 
      <td height="50" colspan="2"> 
        <div align="center"> 
                  <input type="submit" name="Submit" value=" ȷ �� ">
                    <input name="AdminID" value="<% = AdminID %>" type="hidden" id="AdminID">
                    <input name="Result" type="hidden" id="Result" value="Submit">
                    <input type="button" onClick="window.close();" name="Submit2" value=" ȡ �� ">
        </div></td>
    </tr>
 </form>
  </table>
</div>
</body>
</html>
<%
if Result = "Submit" then
	Dim PassWord,AffirmPassWord,RsAdminObj,ReturnCheckInfo
	PassWord = Replace(Replace(Request.Form("PassWord"),"'",""),"""","")
	AffirmPassWord = Replace(Replace(Request.Form("AffirmPassWord"),"'",""),"""","")
	AdminID = Replace(Replace(Request.Form("AdminID"),"'",""),"""","")
	Set RsAdminObj = Server.CreateObject(G_FS_RS)
	RsAdminObj.Open "Select * from FS_Admin where ID="& AdminID &"",Conn
	if RsAdminObj.Eof then
		Set Conn = Nothing
		%>
		<script>alert('�˹���Ա�Ѿ���ɾ��');dialogArguments.location.reload();window.close();</script>
		<%
	end if
	if Len(PassWord) < 6 then
		Set Conn = Nothing
		%>
		<script>alert('��������Ҫ��λ');</script>
		<%
		Response.End  
	end if
	if PassWord <> AffirmPassWord then
		Set Conn = Nothing
		%>
		<script>alert('ȷ�����벻��');</script>
		<%
		Response.End  
	end if
	On Error Resume Next
	Conn.Execute("update FS_Admin set PassWord='" & md5(PassWord,16) & "' where ID=" & AdminID & "")
	if Err.Number = 0 then
		%>
		<script>dialogArguments.location.reload();window.close();</script>
		<%
	else
		%>
		<script>alert('��������');</script>
		<%
	end if
end if
Set Conn = Nothing
%>