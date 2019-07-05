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
<title>复制帮助关键字信息</title>
</head>

<body topmargin="0" leftmargin="0" style="margin:0;overflow-y:auto">
<%
Dim HelpID
HelpID = Request.QueryString("ID")
HelpID = Replace(HelpID,"'","")

If HelpID="" Then Response.write "<script language='javascript'>alert('无效的数据');</script>":Response.end

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