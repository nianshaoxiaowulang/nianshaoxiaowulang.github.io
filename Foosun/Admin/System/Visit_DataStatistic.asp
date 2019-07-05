<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<%
'==============================================================================
'软件名称：风讯网站信息管理系统
'当前版本：Foosun Content Manager System(FoosunCMS V3.1.0930)
'最新更新：2005.10
'==============================================================================
'Copyright (C) 2002-2004 Foosun.Net  All rights reserved.
'商业注册联系：028-85098980-601,项目开发：028-85098980-606、609,客户支持：608
'产品咨询QQ：394226379,159410,125114015
'技术支持QQ：315485710,66252421 
'项目开发QQ：415637671，655071
'程序开发：四川风讯科技发展有限公司(Foosun Inc.)
'Email:service@Foosun.cn
'MSN：skoolls@hotmail.com
'论坛支持：风讯在线论坛(http://bbs.foosun.net)
'官方网站：www.Foosun.cn  演示站点：test.cooin.com 
'网站通系列(智能快速建站系列)：www.ewebs.cn
'==============================================================================
'免费版本请在程序首页保留版权信息，并做上本站LOGO友情连接
'风讯公司保留此程序的法律追究权利
'如需进行2次开发，必须经过风讯公司书面允许。否则将追究法律责任
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
	'转到网站维护页面
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
<title>网站简要信息统计</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script src="../JS/PublicJS.js" language="JavaScript"></script>
<body bgcolor="#FFFFFF" topmargin="2" leftmargin="2">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="28" class="ButtonListLeft"> <div align="center"><strong>网站简要信息统计</strong></div></td>
  </tr>
</table>
<br>
<table width="85%" border="0" align="center" cellpadding="3" cellspacing="1" bordercolor="e6e6e6" bgcolor="dddddd">
  <tr bgcolor="#FFFFFF"> 
    <td width="19%" height="30"> 
      <div align="center">网站名称</div></td>
    <td width="81%"> 
      <% = WebName %></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="30"> 
      <div align="center">管 理 员</div></td>
    <td height="30"> 
      <% = WebAdmin %></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="30"> 
      <div align="center">网站地址</div></td>
    <td height="30"> 
      <% = WebUrl %></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="30"> 
      <div align="center">网站信箱</div></td>
    <td height="30"> 
      <% = WebEmail %></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="30"> 
      <div align="center">网站简介</div></td>
    <td height="30"> 
      <% = WebIntro %></td>
  </tr>
</table>
<p>
<table width="85%" border="0" align="center" cellpadding="3" cellspacing="1" bordercolor="e6e6e6" bgcolor="dddddd">
  <tr bgcolor="#FFFFFF"> 
    <td width="20%"> 
      <div align="center">总 访问人数</div></td>
    <td width="30%"> 
      <% = VisitAllNum %></td>
    <td width="20%"> 
      <div align="center">开始统计日期</div></td>
    <td width="30%"> 
      <% = WebCountTime %></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td> 
      <div align="center">今日 访问量</div></td>
    <td> 
      <% =VisitToday %></td>
    <td> 
      <div align="center">本月 访 问量</div></td>
    <td> 
      <% =VisitMonth %></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td> 
      <div align="center">统 计 天 数</div></td>
    <td> 
      <% =NumofDays %></td>
    <td> 
      <div align="center">平均日访问量</div></td>
    <td> 
      <% = AverageNum%></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td> 
      <div align="center">整站页面刷新</div></td>
    <td> 
      <% =RefreashNum %></td>
    <td> 
      <div align="center">预计本日访问</div></td>
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