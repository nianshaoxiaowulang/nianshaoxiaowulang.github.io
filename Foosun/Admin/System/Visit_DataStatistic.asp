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
if Not JudgePopedomTF(Session("Name"),"P080503") then Call ReturnError1()
Dim Days,Months,Years,TempObj,SunObj,SunNum,VisitTodayNum,VisitMonthNum,VisitAllNums,TempObjs,TeempTimestr
Days = Day(Now())
Months = Month(Now())
Years = Year(Now())
SunNum = 0
Set TempObj = Conn.Execute("Select WebCountTime from FS_WebInfo")
If TempObj.eof then
	TeempTimestr = Now()
Else
	TeempTimestr = TempObj("WebCountTime")
End If
If IsSqlDataBase=0 then
	Set SunObj = Conn.Execute("Select LoginNum from FS_FlowStatistic where VisitTime>#"&TeempTimestr&"#")
Else
	Set SunObj = Conn.Execute("Select LoginNum from FS_FlowStatistic where VisitTime>'"&TeempTimestr&"'")
End If
If Not SunObj.eof then
Do while not SunObj.eof
	SunNum = SunNum + clng(SunObj("LoginNum"))
	SunObj.MoveNext
Loop
End If
SunObj.Close
Set SunObj = Nothing
Set TempObjs = Conn.Execute("Select Count(ID) from FS_FlowStatistic where day(VisitTime) = '"&Days&"' and month(VisitTime)='"&Months&"' and year(VisitTime)='"&Years&"'")
	VisitTodayNum = Clng(TempObjs(0))
Set TempObjs = Conn.Execute("Select Count(ID) from FS_FlowStatistic where month(VisitTime)='"&Months&"' and year(VisitTime)='"&Years&"'")
	VisitMonthNum = Clng(TempObjs(0))
If IsSqlDataBase=0 then
	Set TempObjs = Conn.Execute("Select Count(ID) from FS_FlowStatistic where VisitTime>#"&TeempTimestr&"#")
Else
	Set TempObjs = Conn.Execute("Select Count(ID) from FS_FlowStatistic where VisitTime>'"&TeempTimestr&"'")
End If
	VisitAllNums = Clng(TempObjs(0))
TempObjs.Close
Set TempObjs = Nothing
TempObj.Close
Set TempObj = Nothing
Conn.Execute("Update FS_WebInfo Set VisitToday="&VisitTodayNum&",VisitMonth="&VisitMonthNum&",VisitAllNum="&VisitAllNums&",RefreashNum = "&Clng(SunNum)&"")
Dim Sql,RsWebObj,WebName,WebUrl,WebIntro,WebEmail,WebAdmin,WebCountTime,VisitAllNum,VisitToday,VisitMonth,RefreashNum
Sql ="Select * from FS_WebInfo"
Set RsWebObj = Server.CreateObject(G_FS_RS)
	RsWebObj.Open Sql,Conn,3,3
if Not RsWebObj.Eof then
	WebName = RsWebObj("WebName")
	WebUrl = RsWebObj("WebUrl")
	WebIntro = RsWebObj("WebIntro")
	WebEmail = RsWebObj("WebEmail")
	WebAdmin = RsWebObj("WebAdmin")
	WebCountTime = RsWebObj("WebCountTime")
	VisitAllNum = RsWebObj("VisitAllNum")
	VisitToday = RsWebObj("VisitToday")
	VisitMonth = RsWebObj("VisitMonth")
	RefreashNum = RsWebObj("RefreashNum")
else
	'ת����վά��ҳ��
end if
%>
<%
	Dim ForseeVisitToday,I,NumofDays,AverageNum,Tnum,TNumStr
	NumofDays=DATEDIFF("d",WebCountTime,Date())
	If NumofDays = 0 then
		NumofDays = 1
	End If
	AverageNum=CLng(VisitAllNum/NumofDays*1000)/1000
	I= Now()-Date()
	ForseeVisitToday = Round(VisitToday/I*(1-I) + VisitToday)
%>
<html>
<head>
<title>��վ��Ҫ��Ϣͳ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script src="../JS/PublicJS.js" language="JavaScript"></script>
<body bgcolor="#FFFFFF" topmargin="2" leftmargin="2">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="28" class="ButtonListLeft"> <div align="center"><strong>��վ��Ҫ��Ϣͳ��</strong></div></td>
  </tr>
</table>
<br>
<table width="85%" border="0" align="center" cellpadding="3" cellspacing="1" bordercolor="e6e6e6" bgcolor="dddddd">
  <tr bgcolor="#FFFFFF"> 
    <td width="19%" height="30"> 
      <div align="center">��վ����</div></td>
    <td width="81%"> 
      <% = WebName %></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="30"> 
      <div align="center">�� �� Ա</div></td>
    <td height="30"> 
      <% = WebAdmin %></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="30"> 
      <div align="center">��վ��ַ</div></td>
    <td height="30"> 
      <% = WebUrl %></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="30"> 
      <div align="center">��վ����</div></td>
    <td height="30"> 
      <% = WebEmail %></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="30"> 
      <div align="center">��վ���</div></td>
    <td height="30"> 
      <% = WebIntro %></td>
  </tr>
</table>
<p>
<table width="85%" border="0" align="center" cellpadding="3" cellspacing="1" bordercolor="e6e6e6" bgcolor="dddddd">
  <tr bgcolor="#FFFFFF"> 
    <td width="20%"> 
      <div align="center">�� ��������</div></td>
    <td width="30%"> 
      <% = VisitAllNum %></td>
    <td width="20%"> 
      <div align="center">��ʼͳ������</div></td>
    <td width="30%"> 
      <% = WebCountTime %></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td> 
      <div align="center">���� ������</div></td>
    <td> 
      <% =VisitToday %></td>
    <td> 
      <div align="center">���� �� ����</div></td>
    <td> 
      <% =VisitMonth %></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td> 
      <div align="center">ͳ �� �� ��</div></td>
    <td> 
      <% =NumofDays %></td>
    <td> 
      <div align="center">ƽ���շ�����</div></td>
    <td> 
      <% = AverageNum%></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td> 
      <div align="center">��վҳ��ˢ��</div></td>
    <td> 
      <% =RefreashNum %></td>
    <td> 
      <div align="center">Ԥ�Ʊ��շ���</div></td>
    <td> 
      <% =ForseeVisitToday %></td>
  </tr>
</table>
</body>
</html>
<%
	RsWebObj.Close
	Conn.Close
	Set RsWebObj=nothing
	Set Conn=nothing
%>