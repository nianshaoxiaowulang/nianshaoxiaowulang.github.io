<% Option Explicit %>
<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Const.asp" -->
<%
Dim DBC,Conn,HelpConn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
DBC.ConnStr = "DBQ=" + Server.MapPath("Foosun_help.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
Set HelpConn = DBC.OpenConnection()
Set DBC = Nothing
'==============================================================================
'������ƣ�FoosunHelp System Form FoosunCMS
'��ǰ�汾��Foosun Content Manager System 3.0 ϵ��
'���¸��£�2005.12
'==============================================================================
'��ҵע����ϵ��028-85098980-601,602 ����֧�֣�028-85098980-605��607,�ͻ�֧�֣�608
'��Ʒ��ѯQQ��159410,394226379,125114015,655071
'����֧��:���г���ʹ�����⣬�����ʵ�bbs.foosun.net���ǽ���ʱ�ش���
'���򿪷�����Ѷ������ & ��Ѷ���������
'��̳֧�֣���Ѷ������̳(http://bbs.foosun.net)
'�ٷ���վ��www.Foosun.net  ��ʾվ�㣺test.cooin.com    
'��վ����ר����www.cooin.com
'==============================================================================
'��Ѱ汾����������ҳ������Ȩ��Ϣ�������ϱ�վLOGO��������
'==============================================================================
%>
<!--#include file="../../Inc/Session.asp" -->
<!--#include file="../../Inc/CheckPopedom.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../CSS/FS_css.css" rel="stylesheet" type="text/css">
<title>ɾ�������ؼ�����Ϣ</title>
</head>

<body topmargin="0" leftmargin="0" style="margin:0;overflow-y:auto">
<%
if Not JudgePopedomTF(Session("Name"),"P070803") then Call ReturnError1()
Dim ID
ID = Request.QueryString("ID")
ID = Replace(ID,"'","")
If ID=0 Then Response.write "<script language='javascript'>alert('��Ч������');</script>":Response.end
HelpConn.Execute("Delete From [FS_Help] where ID in ("&ID&")")
Response.write "<script language='javascript'>parent.location.reload();</script>"
Set Conn = Nothing
Set HelpConn=Nothing
%>
</body>
</html>