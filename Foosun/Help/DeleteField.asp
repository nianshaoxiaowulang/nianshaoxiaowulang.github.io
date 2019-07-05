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
'软件名称：FoosunHelp System Form FoosunCMS
'当前版本：Foosun Content Manager System 3.0 系列
'最新更新：2005.12
'==============================================================================
'商业注册联系：028-85098980-601,602 技术支持：028-85098980-605、607,客户支持：608
'产品咨询QQ：159410,394226379,125114015,655071
'技术支持:所有程序使用问题，请提问到bbs.foosun.net我们将及时回答您
'程序开发：风讯开发组 & 风讯插件开发组
'论坛支持：风讯在线论坛(http://bbs.foosun.net)
'官方网站：www.Foosun.net  演示站点：test.cooin.com    
'网站建设专区：www.cooin.com
'==============================================================================
'免费版本请在新闻首页保留版权信息，并做上本站LOGO友情连接
'==============================================================================
%>
<!--#include file="../../Inc/Session.asp" -->
<!--#include file="../../Inc/CheckPopedom.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../CSS/FS_css.css" rel="stylesheet" type="text/css">
<title>删除帮助关键字信息</title>
</head>

<body topmargin="0" leftmargin="0" style="margin:0;overflow-y:auto">
<%
if Not JudgePopedomTF(Session("Name"),"P070803") then Call ReturnError1()
Dim ID
ID = Request.QueryString("ID")
ID = Replace(ID,"'","")
If ID=0 Then Response.write "<script language='javascript'>alert('无效的数据');</script>":Response.end
HelpConn.Execute("Delete From [FS_Help] where ID in ("&ID&")")
Response.write "<script language='javascript'>parent.location.reload();</script>"
Set Conn = Nothing
Set HelpConn=Nothing
%>
</body>
</html>