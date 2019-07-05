<% Option Explicit %>
<!--#include file="../../Inc/Const.asp" -->
<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Function.asp" -->
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
<!--#include file="../../Inc/Session.asp" -->
<!--#include file="../../Inc/CheckPopedom.asp" -->
<%
Dim RsMenuConfigObj,HaveValueTF
Set RsMenuConfigObj = Conn.execute("Select IsShop From FS_Config")
if RsMenuConfigObj("IsShop") = 1 then
	HaveValueTF = True
Else
	HaveValueTF = False
End if
Set RsMenuConfigObj = Nothing
%><html>
<head>
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<meta name="GENERATOR" Content="Microsoft Visual Studio .NET 7.1">
<meta http-equiv="pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../CSS/FS_css.css" rel="stylesheet">
</head>
<script language="JavaScript">

function StartEnlarge(e)
{
	top.StartEnlarge(e);
}

function StartShrink(e)
{
	top.StartShrink(e);
}

function ShrinkFrame(e)
{
	top.ShrinkFrame(e);
}

function ShowDeskTop()
{
	top.GetNavFoldersObject().location='ShortCutPage.asp';
}
</script>
<script language="JavaScript" src="../SysJS/PublicJS.js"></script>
<!--不允许修改，否则给你急  :)-->
<body topmargin="2" leftmargin="2" class="FolderToolbar" oncontextmenu="return false;" onmouseout="StartShrink(event);" onmouseover="StartEnlarge(event);">
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#999999" class="FolderToolbar" ID="Table1">
  <tr bgcolor="#EEEEEE">
    <td height="26">
		<table width="100%" height="20" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="12"><img alt="打开快捷菜单" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" id="DeskTopImg" border="0" src="Images/cmsv31_show.png" width="12" height="12" onclick="ShowDeskTop();return false;" class="BtnMouseOut"></td>
          <td width=30 id="RightToolbarContainer" align="center" alt="生成首页" onClick="window.open('Refresh/RefreshIndex.asp','fs_main')" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">首页</td>
          <td width=7 class="Gray">|</td>
          <td width=30 align="center" alt="生成栏目" onClick="window.open('Refresh/RefreshClass.asp','fs_main')" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">栏目</td>
          <td width=6 class="Gray">|</td>
          <td width=31 align="center" alt="生成新闻" onClick="window.open('Refresh/RefreshNews.asp','fs_main')" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">新闻</td>
          <%If HaveValueTF = True then%>
		  <td width="9" align="right" valign="middle" class="Gray">|</td>
          <td width="34" align="right" valign="middle"onClick="window.open('Refresh/Mall_Refresh.asp','fs_main')" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut"><div align="left">商城</div></td>
		  <%End if%>
          <td align="right" valign="middle"><img alt="隐藏菜单" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" id="CancelImg" border="0" src="Images/cmsv31_close.png" onclick="ShrinkFrame();return false;" class="BtnMouseOut">&nbsp;</td>
        </tr>
      </table> </td>
</tr>
</table>
</body>
</html>
