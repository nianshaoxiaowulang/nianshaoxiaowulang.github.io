<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../Refresh/Function.asp" -->
<!--#include file="../Refresh/RefreshFunction.asp" -->
<!--#include file="../Refresh/SelectFunction.asp" -->
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
if Not JudgePopedomTF(Session("Name"),"P010000") then Call ReturnError()
Dim NewsID,DownLoadID,Action,ProductID
NewsID = Request("NewsID")
DownLoadID = Request("DownLoadID")
ProductID=Request("ProductID")
Action = Request("Action")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>生成新闻</title>
</head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet">
<body topmargin="0" leftmargin="0">
<table width="100%" border="0" cellspacing="8" cellpadding="0">
  <form name="DelForm" method="post" action="">
    <tr> 
      <td width="21%"> <div align="center"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
      <td width="79%" colspan="2"> 确定要生成吗? 
        <input name="Action" type="hidden" id="Action" value="Submit"> 
        <input name="NewsID" type="hidden" id="NewsID" value="<% = NewsID %>"> 
        <input name="DownLoadID" type="hidden" id="DownLoadID" value="<% = DownLoadID %>"></td>
    </tr>
    <tr> 
      <td colspan="3"> <div align="center"> 
          <input name="Submitsadf" type="submit" id="Submitsadf" value=" 生 成 ">
          
          <input type="button" onClick="window.close();" name="Submit3" value=" 取 消 ">
        </div></td>
    </tr>
  </form>
</table>
</body>
</html>
<%
if Action = "Submit" then
	Dim RsRefreshObj,RefreshNewsNum,RefreshDownLoadNum,RefreshProductNum,AlertStr
	RefreshNewsNum = 0
	RefreshDownLoadNum = 0
	if NewsID <> "" then
		if Not JudgePopedomTF(Session("Name"),"P010510") then Call ReturnError()
		NewsID = Replace(NewsID,"***","','")
		Set RsRefreshObj = Conn.Execute("Select * from FS_News where AuditTF=1 and NewsID in ('" & NewsID & "')")
		do while Not RsRefreshObj.Eof
			RefreshNews RsRefreshObj
			RefreshNewsNum = RefreshNewsNum + 1
			RsRefreshObj.MoveNext
		Loop
		Set RsRefreshObj = Nothing
			AlertStr = "生成" & RefreshNewsNum & "条新闻"
	end if
	if DownLoadID <> "" then
		if Not JudgePopedomTF(Session("Name"),"P010706") then Call ReturnError()
		DownLoadID = Replace(DownLoadID,"***","','")
		Set RsRefreshObj = Conn.Execute("Select * from FS_DownLoad where AuditTF=1 and DownLoadID in ('" & DownLoadID & "')")
		do while Not RsRefreshObj.Eof
			RefreshDownLoad RsRefreshObj
			RefreshDownLoadNum = RefreshDownLoadNum + 1
			RsRefreshObj.MoveNext
		Loop
		Set RsRefreshObj = Nothing
		AlertStr = "生成" & RefreshDownLoadNum & "条下载"
	end if
	if ProductID <> "" then
		if Not JudgePopedomTF(Session("Name"),"P010806") then Call ReturnError()
		ProductID = Replace(ProductID,"***",",")
		Set RsRefreshObj = Conn.Execute("Select * from FS_Shop_Products where id in (" & ProductID & ")")

		do while Not RsRefreshObj.Eof
			RefreshOneMallProduct RsRefreshObj
			RefreshProductNum = RefreshProductNum + 1
			RsRefreshObj.MoveNext
		Loop
		Set RsRefreshObj = Nothing
		AlertStr = "生成" & RefreshProductNum & "个商品"
	end if

	Response.Write("<script language=""JavaScript"">alert('" & AlertStr & "');window.close();</script>")
end if
%>