<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
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

%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P070403") then Call ReturnError()
Dim Result,TypeStr
Dim OrdinaryID,OperateType,Sql
OrdinaryID = Request("OrdinaryID")
OperateType = Request("OperateType")
Result = Request.Form("Result")
Select Case Cint(OperateType)
  Case "1" TypeStr="�ؼ���"
  Case "2" TypeStr="��Դ"
  Case "3" TypeStr="����"
  Case "4" TypeStr="�༭"
  Case "5" TypeStr="�ڲ�����"
End Select
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�������ɾ��</title>
</head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet">
<body topmargin="0" leftmargin="0">
<table width="100%" border="0" align="center">
  <form name="" action="" method="post">
  <tr>
    <td width="32%"><div align="center"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
    <td width="68%">ȷ��Ҫɾ��<%=TypeStr%>��</td>
  </tr>
  <tr>
    <td colspan="2"><div align="center">
        <input name="Submit" type="submit" id="Submit" value=" ȷ �� ">
        <input name="OperateType" value="<% = OperateType %>" type="hidden" id="OperateType">
        <input name="Result" type="hidden" id="Result" value="Submit">
        <input type="hidden" value="<% = OrdinaryID %>" name="OrdinaryID">
        <input name="Submit1"  onClick="window.close();"type="reset" id="Submit1" value=" ȡ �� ">
    </div></td>
  </tr>
  </form>
</table>
</body>
</html>
<%
if Result = "Submit" then
	if OrdinaryID <> "" then
		Sql = "Delete from FS_Routine where ID in (" & Replace(OrdinaryID,"***",",") & ") and Type=" & OperateType
		Conn.Execute(Sql)
		Set Conn = Nothing
	end if
	if Err.Number = 0 then
		%>
		<script language="JavaScript">
			dialogArguments.location.reload();
			window.close();
		</script>
		<%
	else
		%>
		<script language="JavaScript">
			alert('ɾ��ʧ��');
			dialogArguments.location.reload();
			window.close();
		</script>
		<%
	end if
end if
Set Conn = Nothing
%>