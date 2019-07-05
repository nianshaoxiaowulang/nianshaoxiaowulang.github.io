<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="inc/Config.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
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
'判断权限
if Not JudgePopedomTF(Session("Name"),"P080300") then Call ReturnError1()
'判断权限结束
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
<title>设置新闻属性</title>
</head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet">
<body topmargin="0" leftmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
 <form name="SetForm" action="" method="post">
  <tr> 
    <td width="100" rowspan="3"> 
      <div align="center"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
    <td height="26"><div align="center">选择设置的新闻属性
          <input type="hidden" name="NewsIDStr" value="<% = Request("NewsIDStr") %>">
          <input type="hidden" name="Action" value="Submit">
        </div></td>
  </tr>
  <tr> 
    <td height="36"> 
      <div align="center"> 
        <input name="PicNews" type="checkbox" id="PicNews" value="1">
        图片新闻 
        <input name="RecTF" type="checkbox" id="RecTF" value="1">
        推荐新闻 
        <input name="TodayNewsTF" type="checkbox" id="TodayNewsTF" value="1">
        今日头条</div></td>
  </tr>
  <tr> 
    <td height="36"> 
      <div align="center"> 
        <input name="MarqueeNews" type="checkbox" id="MarqueeNews" value="1">
        滚动新闻 
        <input name="SBSNews" type="checkbox" id="SBSNews" value="1">
        并排新闻 
        <input name="ReviewTF" type="checkbox" id="ReviewTF" value="1">
        允许评论</div></td>
  </tr>
  <tr> 
    <td height="46" colspan="2">
<div align="center"> 
          <input name="Submitfgsfd" type="submit" id="Submitfgsfd" value=" 确 定 ">
        &nbsp;&nbsp;&nbsp;&nbsp; 
        <input name="Submit2fasd" type="button" id="Submit2fasd" onClick="window.close();" value=" 取 消 ">
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