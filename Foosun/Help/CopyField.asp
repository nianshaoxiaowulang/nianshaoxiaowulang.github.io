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
<title>���ư����ؼ�����Ϣ</title>
</head>

<body topmargin="0" leftmargin="0" style="margin:0;overflow-y:auto">
<%
Dim HelpID
HelpID = Request.QueryString("ID")
HelpID = Replace(HelpID,"'","")

If HelpID="" Then Response.write "<script language='javascript'>alert('��Ч������');</script>":Response.end

Dim FuncName,FileName,PageField,HelpContent,HelpSingleContent

Dim tempRs
Set tempRs = Server.CreateObject(G_FS_RS)
dim i,IsFind
HelpID = split(HElpID,",")

for i=Lbound(HelpID) to Ubound(HelpID)
	tempRs.open "Select * From [Fs_Help] where id="&Clng(HelpID(i)),HelpConn,1,1
	IsFind=False
	if not tempRs.eof then
		FuncName = tempRs("FuncName")
		FileName = tempRs("FileName")
		PageField = tempRs("PageField")
		HelpContent = tempRs("HelpContent")
		HelpSingleContent = tempRs("HelpSingleContent")
		IsFind = true
	end if
	tempRs.close
	If IsFind Then
		tempRs.open "Select * From [Fs_Help]",HelpConn,1,3
		tempRs.addnew
		tempRs("FuncName") = FuncName
		tempRs("FileName") = FileName
		tempRs("PageField") = PageField
		tempRs("HelpContent") = HelpContent
		tempRs("HelpSingleContent") = HelpSingleContent
		tempRs("SvTime") = now
		tempRs.update
		tempRs.close
	End If
Next
Response.write "<script language='javascript'>parent.location.reload();</script>"

set tempRs = Nothing


Set Conn = Nothing
Set HelpConn = Nothing
%>
</body>
</html>