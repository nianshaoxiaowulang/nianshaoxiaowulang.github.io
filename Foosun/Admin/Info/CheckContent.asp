<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
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
if Not JudgePopedomTF(Session("Name"),"P010000") then Call ReturnError()
Dim NewsID,DownLoadID,OperateType,TempStr,NewsOrDown
OperateType = Request("OperateType")
DownLoadID = Request("DownLoadID")
NewsID = Cstr(Request("NewsID"))
If NewsID<>"" then 
	if Not JudgePopedomTF(Session("Name"),"P010504") then Call ReturnError()
	NewsOrDown="����"
End If
If DownLoadID<>"" then
	if Not JudgePopedomTF(Session("Name"),"P010703") then Call ReturnError()
	NewsOrDown="����"
End If
if NewsID <> "" OR DownLoadID <> "" then
	if OperateType = "UnCheck" then
		TempStr = "�������"
	else
		TempStr = "���"
	end if
Else
	Response.Write("<script>alert(""�������ݴ���"");window.close();</script>")
	response.end
end if 
%>
<html>
<head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�������</title>
</head>
<body leftmargin="0" topmargin="0">
<form action="" name="JSDellForm" method="post">
<table width="100%" border="0" cellspacing="5" cellpadding="0">
  <tr> 
    <td width="7%" height="10">&nbsp;</td>
    <td width="28%" rowspan="3"><div align="center"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
    <td width="59%">&nbsp;</td>
    <td width="6%" height="10">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>��ȷ��Ҫ<%=TempStr%><%=NewsOrDown%>?</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td height="2">&nbsp;</td>
    <td height="2">&nbsp;</td>
    <td height="2">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td colspan="2"><div align="center"> 
        <input type="submit" name="Submit" value=" ȷ �� ">
        <input type="hidden" name="action" value="trues">
        <input type="button" name="Submit2" value=" ȡ �� " onClick="window.close();">
      </div></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td height="10">&nbsp;</td>
    <td height="10" colspan="2">&nbsp;</td>
    <td height="10">&nbsp;</td>
  </tr>
</table>
</form>
</body>
</html>
<%
if Request.Form("action") = "trues" then
	if NewsID <> "" Then
	NewsID = Replace(Replace(Replace(Replace(Replace(NewsID,"'",""),"and",""),"select",""),"or",""),"union","")
		NewsID = Replace(NewsID,"***","','")
		if OperateType = "UnCheck" then
			Conn.Execute("Update FS_News Set AuditTF=0 where NewsID in ('" & NewsID & "')")
		elseif  OperateType = "Check" then
		'response.write "Update News Set AuditTF=1 where NewsID in ('" & NewsID & "')"
		'response.end
			Conn.Execute("Update FS_News Set AuditTF=1 where NewsID in ('" & NewsID & "')")
		End if
	end if
	if DownLoadID <> "" Then
		DownLoadID = Replace(Replace(Replace(Replace(Replace(DownLoadID,"'",""),"and",""),"select",""),"or",""),"union","")
		DownLoadID = Replace(DownLoadID,"***","','")
		if OperateType = "UnCheck" then
			Conn.Execute("Update FS_DownLoad Set AuditTF=0 where DownLoadID in ('" & DownLoadID & "')")
		elseif  OperateType = "Check" then
			Conn.Execute("Update FS_DownLoad Set AuditTF=1 where DownLoadID in ('" & DownLoadID & "')")
		End if
	end if
	Response.Write("<script>window.close();</script>")
end if
%>