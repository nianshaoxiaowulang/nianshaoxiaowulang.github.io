<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="inc/Config.asp" -->
<!--#include file="inc/Function.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
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
Dim DBC,Conn,CollectConn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
DBC.ConnStr = CollectDBConnectionStr
Set CollectConn = DBC.OpenConnection()
Set DBC = Nothing
'�ж�Ȩ��
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P080303") then Call ReturnError()
'�ж�Ȩ�޽���
Sub ShowInfo(InfoStr)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ƶ�����</title>
</head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet">
<body topmargin="0" leftmargin="0" scroll=no>
<div align="center">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="120">
<div align="center"><img src="../../Images/Info.gif" width="34" height="33"></div></td>
      <td height="60"> <div align="left"> 
          <% = InfoStr %>
        </div></td>
    </tr>
    <tr> 
      <td colspan="2"> <div align="center"> 
          <input onClick="dialogArguments.location.reload();window.close();" type="button" name="Submit2" value=" ȷ �� ">
        </div></td>
    </tr>
  </table>
</div>
</body>
</html>
<%
End Sub

Dim MoveNewsList,RsSysNewsObj,SysNewsSql,MoveNewsNum
Dim SysTemplet,RsTempObj,Sql,RsNewsObj,AuditTF
MoveNewsNum = 0
MoveNewsList = Replace(Request("NewsIDStr"),"***",",")
if MoveNewsList = "" then
	ShowInfo "û��ѡ������"
	Response.End
end if
if MoveNewsList = "All" then
	Sql = "Select * from FS_News where History=0"
else
	Sql = "Select * from FS_News where ID in (" & MoveNewsList & ")"
end if
'On Error Resume Next
Set RsNewsObj = CollectConn.Execute(Sql)
Dim NameRule
if Not RsNewsObj.Eof then
	NameRule=conn.execute("select NewsFileName from FS_Config")(0)
	SysNewsSql = "Select * from FS_News where 1=0"
	Set RsSysNewsObj = Server.CreateObject("ADODB.RecordSet")
	RsSysNewsObj.Open SysNewsSql,Conn,3,3
	do while Not RsNewsObj.Eof
		Set RsTempObj = CollectConn.Execute("Select SysTemplet,Audit from FS_Site where ID=" & RsNewsObj("SiteID"))
		if Not RsTempObj.Eof then
			SysTemplet = RsTempObj("SysTemplet")
			AuditTF = RsTempObj("Audit")
		else
			SysTemplet = "/Templets/NewsClass/News.htm"
			AuditTF = False
		end if
		Set RsTempObj = Nothing
		RsSysNewsObj.AddNew
		RsSysNewsObj("NewsID") = GetRandomID18
		RsSysNewsObj("Title") = RsNewsObj("Title")
		RsSysNewsObj("TitleStyle") = "#UUUUUU00"
		RsSysNewsObj("ClassID") = RsNewsObj("ClassID")
		RsSysNewsObj("Content") = RsNewsObj("Content")
		RsSysNewsObj("NewsTemplet") = SysTemplet
		RsSysNewsObj("FileName") = NewsFileName(NameRule,RsNewsObj("ClassID"),RsSysNewsObj("NewsID"))
		RsSysNewsObj("FileExtName") = "htm"
		RsSysNewsObj("Path") = "/" & year(now())&"-"&month(now())&"/"&day(now())             '����·�� 
		RsSysNewsObj("AddDate") = RsNewsObj("AddDate")
		RsSysNewsObj("Author") = RsNewsObj("Author")
		RsSysNewsObj("TxtSource") = RsNewsObj("Source")
		if RsNewsObj("PicNews") = True then
			RsSysNewsObj("PicNewsTF") = 1
		else
			RsSysNewsObj("PicNewsTF") = 0
		end if
		if RsNewsObj("RecTF") = True then
			RsSysNewsObj("RecTF") = 1
		else
			RsSysNewsObj("RecTF") = 0
		end if
		if RsNewsObj("TodayNewsTF") = True then
			RsSysNewsObj("TodayNewsTF") = 1
		else
			RsSysNewsObj("TodayNewsTF") = 0
		end if
		if RsNewsObj("MarqueeNews") = True then
			RsSysNewsObj("MarqueeNews") = 1
		else
			RsSysNewsObj("MarqueeNews") = 0
		end if
		if RsNewsObj("SBSNews") = True then
			RsSysNewsObj("SBSNews") = 1
		else
			RsSysNewsObj("SBSNews") = 0
		end if
		if RsNewsObj("ReviewTF") = True then
			RsSysNewsObj("ReviewTF") = 1
		else
			RsSysNewsObj("ReviewTF") = 0
		end if
		if AuditTF = True then
			RsSysNewsObj("AuditTF") = 1
		else
			RsSysNewsObj("AuditTF") = 0
		end if
		MoveNewsNum = MoveNewsNum + 1
		RsNewsObj.MoveNext
	loop
	RsSysNewsObj.UpDate
	RsSysNewsObj.Close
	Set RsSysNewsObj = Nothing
end if
Set RsNewsObj = Nothing
if Request("DelNews") = "true" then
	Sql = "Delete from FS_News where ID in (" & MoveNewsList & ")"
else
	if MoveNewsList = "All" then
		Sql = "Update FS_News Set History=1"
	else
		Sql = "Update FS_News Set History=1 where ID in (" & MoveNewsList & ")"
	end if
end if
CollectConn.Execute(Sql)
if Err.Number = 0 then
	ShowInfo "ת�Ƴɹ�" & MoveNewsNum & "������"
else
	ShowInfo "ת��ʧ��"
end if
Set CollectConn = Nothing
Set Conn = Nothing
%>