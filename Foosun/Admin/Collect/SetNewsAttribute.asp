<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="inc/Config.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
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
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
'�ж�Ȩ��
if Not JudgePopedomTF(Session("Name"),"P080300") then Call ReturnError1()
'�ж�Ȩ�޽���
Dim Action,NewsIDStr,PicNews,RecTF,TodayNewsTF,MarqueeNews,SBSNews,ReviewTF,Sql
Action = Request("Action")
if Action = "Submit" then
	NewsIDStr = Request("NewsIDStr")
	if NewsIDStr <> "" then
		NewsIDStr = Replace(NewsIDStr,"***",",")
		PicNews = Request("PicNews")
		if PicNews = "1" then
			PicNews = 1
		else
			PicNews = 0
		end if
		CollectConn.Execute("Update FS_News set PicNews=" & PicNews & " where ID in (" & NewsIDStr & ")")
		RecTF = Request("RecTF")
		if RecTF = "1" then
			RecTF = 1
		else
			RecTF = 0
		end if
		CollectConn.Execute("Update FS_News set RecTF=" & RecTF & " where ID in (" & NewsIDStr & ")")
		TodayNewsTF = Request("TodayNewsTF")
		if TodayNewsTF = "1" then
			TodayNewsTF = 1
		else
			TodayNewsTF = 0
		end if
		CollectConn.Execute("Update FS_News set TodayNewsTF=" & TodayNewsTF & " where ID in (" & NewsIDStr & ")")
		MarqueeNews = Request("MarqueeNews")
		if MarqueeNews = "1" then
			MarqueeNews = 1
		else
			MarqueeNews = 0
		end if
		CollectConn.Execute("Update FS_News set MarqueeNews=" & MarqueeNews & " where ID in (" & NewsIDStr & ")")
		SBSNews = Request("SBSNews")
		if SBSNews = "1" then
			SBSNews = 1
		else
			SBSNews = 0
		end if
		CollectConn.Execute("Update FS_News set SBSNews=" & SBSNews & " where ID in (" & NewsIDStr & ")")
		ReviewTF = Request("ReviewTF")
		if ReviewTF = "1" then
			ReviewTF = 1
		else
			ReviewTF = 0
		end if
		CollectConn.Execute("Update FS_News set ReviewTF=" & ReviewTF & " where ID in (" & NewsIDStr & ")")
	end if
	Set Conn = Nothing
	Set CollectConn = Nothing
	Response.Write("<script>dialogArguments.location.reload();window.close();</script>")
	Response.End
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>������������</title>
</head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet">
<body topmargin="0" leftmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
 <form name="SetForm" action="" method="post">
  <tr> 
    <td width="100" rowspan="3"> 
      <div align="center"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
    <td height="26"><div align="center">ѡ�����õ���������
          <input type="hidden" name="NewsIDStr" value="<% = Request("NewsIDStr") %>">
          <input type="hidden" name="Action" value="Submit">
        </div></td>
  </tr>
  <tr> 
    <td height="36"> 
      <div align="center"> 
        <input name="PicNews" type="checkbox" id="PicNews" value="1">
        ͼƬ���� 
        <input name="RecTF" type="checkbox" id="RecTF" value="1">
        �Ƽ����� 
        <input name="TodayNewsTF" type="checkbox" id="TodayNewsTF" value="1">
        ����ͷ��</div></td>
  </tr>
  <tr> 
    <td height="36"> 
      <div align="center"> 
        <input name="MarqueeNews" type="checkbox" id="MarqueeNews" value="1">
        �������� 
        <input name="SBSNews" type="checkbox" id="SBSNews" value="1">
        �������� 
        <input name="ReviewTF" type="checkbox" id="ReviewTF" value="1">
        ��������</div></td>
  </tr>
  <tr> 
    <td height="46" colspan="2">
<div align="center"> 
          <input name="Submitfgsfd" type="submit" id="Submitfgsfd" value=" ȷ �� ">
        &nbsp;&nbsp;&nbsp;&nbsp; 
        <input name="Submit2fasd" type="button" id="Submit2fasd" onClick="window.close();" value=" ȡ �� ">
      </div></td>
  </tr>
 </form>
</table>
</body>
</html>
<%
Set Conn = Nothing
Set CollectConn = Nothing
%>