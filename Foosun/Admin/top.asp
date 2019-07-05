<% Option Explicit %>
<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Const.asp" -->
<%
'==============================================================================
'软件名称：风讯网站信息管理系统
'当前版本：Foosun Content Manager System v3.1 
'最新更新：2004.12
'==============================================================================
'商业注册联系：028-85098980-601,602 技术支持：028-85098980-606、607,客户支持：608
'产品咨询QQ：159410,655071,66252421
'技术支持:所有程序使用问题，请提问到bbs.foosun.net我们将及时回答您
'程序开发：风讯开发组 & 风讯插件开发组
'论坛支持：风讯在线论坛(http://bbs.foosun.net)
'官方网站：www.Foosun.net  演示站点：test.cooin.com    
'网站建设专区：www.cooin.com
'==============================================================================
'免费版本请在新闻首页保留版权信息，并做上本站LOGO友情连接
'==============================================================================
Dim DBC,Conn,URLS
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
Dim RsMenuConfigObj,HaveValueTF
Set RsMenuConfigObj = Conn.execute("Select IsShop From FS_Config")
if RsMenuConfigObj("IsShop") = 1 then
	HaveValueTF = True
Else
	HaveValueTF = False
End if
Set RsMenuConfigObj = Nothing
%>
<!--#include file="../../Inc/Session.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<link href="../../Css/Style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../SysJS/PublicJS.js"></script>
</head>

<body onselectstart="return false;" oncontextmenu="return false;">
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr valign="top">
    <td width="267" background="images/cmsv31_02.png"><img src="images/cmsv31_01.png" width="267" height="60" alt=""></td><td width="54" align="right" background="images/cmsv31_02.png">&nbsp;</td><td background="images/cmsv31_02.png"><a href="Menu_Folders.asp?Action=ContentTree" target="nav_folder_area"><img alt="信息管理" src="images/icon_1.png" width="54" border="0" style="cursor:hand"></a><a href="Menu_Folders.asp?Action=Special" target="nav_folder_area"><img alt="专题管理" src="images/icon_2.png" width="54" border="0" style="cursor:hand"></a><a href="Menu_Folders.asp?Action=NetStation" target="nav_folder_area"><img alt="站点管理" src="images/icon_3.png" width="54" border="0" style="cursor:hand"></a><a href="Menu_Folders.asp?Action=System" target="nav_folder_area"><img alt="系统管理" src="images/icon_4.png" width="54" height="60" border="0" style="cursor:hand"></a><a href="Menu_Folders.asp?Action=UpLoad" target="nav_folder_area"><img alt="虚拟目录管理" src="images/icon_5.png" width="54" height="60" border="0" style="cursor:hand"></a><a href="Menu_Folders.asp?Action=JSManage" target="nav_folder_area"><img alt="JS管理" src="images/icon_6.png" width="54" height="60" border="0" style="cursor:hand"></a><a href="Menu_Folders.asp?Action=OrdinaryManage" target="nav_folder_area"><img alt="常规管理" src="images/icon_7.png" width="54" height="60" border="0" style="cursor:hand"></a><a href="System/ChangePwd.asp" target="fs_main"><img alt="修改管理员密码" src="images/icon_9.png" width="54" height="60" border="0" style="cursor:hand"></a><a href="LoginOut.asp" target="_top"><img alt="退出系统" src="images/icon_a.png" width="54" height="60" border="0" style="cursor:hand"></a> 
    </td>
  </tr>
</table>
</body>
</html>
