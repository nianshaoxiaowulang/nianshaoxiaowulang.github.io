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
if Not JudgePopedomTF(Session("Name"),"P010604") then Call ReturnError()
    Dim NewsID,NewsObj
	If Request("NewsID")<>"" then
		NewsID = Request("NewsID")
	Else
	   Response.Write("<script>alert(""参数传递错误"");dialogArguments.location.reload();window.close();</script>")
	   Response.End
	End If
%>
<html>
<head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>稿件删除</title>
</head>
<body>
<table width="100%" border="0" cellspacing="5" cellpadding="0">
 <form action="" name="JSDellForm" method="post">
  <tr> 
    <td width="6%" height="10">&nbsp;</td>
    <td width="22%" rowspan="3"><div align="center"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
    <td width="72%" height="10">&nbsp;</td>
    </tr>
  <tr> 
    <td>&nbsp;</td>
      <td>您确定要删除稿件?</td>
    </tr>
  <tr>
    <td height="2">&nbsp;</td>
    <td height="2">&nbsp;</td>
    </tr>
  <tr> 
    <td>&nbsp;</td>
    <td colspan="2"><div align="center"> 
        <input type="submit" name="Submit" value=" 确 定 ">
        <input type="hidden" name="action" value="trues">
        <input type="button" name="Submit2" value=" 取 消 " onClick="window.close();">
      </div></td>
    </tr>
 </form>
</table>
</body>
</html>
<%
 If Request.Form("action")="trues" then
 	Dim DCArray,DC_i
	DCArray = Array("")
	DCArray = Split(NewsID,"***")
	For DC_i = 0 to UBound(DCArray)
		Conn.Execute("delete from FS_Contribution where ContID='"&DCArray(DC_i)&"'")
	Next
	Response.write("<script>dialogArguments.location.reload();window.close();</script>")
 	Response.End
 End If
%>