<% Option Explicit %>
<!--#include file="../../Inc/Const.asp" -->
<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Function.asp" -->
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
<!--�������޸ģ�������㼱  :)-->
<body topmargin="2" leftmargin="2" class="FolderToolbar" oncontextmenu="return false;" onmouseout="StartShrink(event);" onmouseover="StartEnlarge(event);">
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#999999" class="FolderToolbar" ID="Table1">
  <tr bgcolor="#EEEEEE">
    <td height="26">
		<table width="100%" height="20" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="12"><img alt="�򿪿�ݲ˵�" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" id="DeskTopImg" border="0" src="Images/cmsv31_show.png" width="12" height="12" onclick="ShowDeskTop();return false;" class="BtnMouseOut"></td>
          <td width=30 id="RightToolbarContainer" align="center" alt="������ҳ" onClick="window.open('Refresh/RefreshIndex.asp','fs_main')" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">��ҳ</td>
          <td width=7 class="Gray">|</td>
          <td width=30 align="center" alt="������Ŀ" onClick="window.open('Refresh/RefreshClass.asp','fs_main')" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">��Ŀ</td>
          <td width=6 class="Gray">|</td>
          <td width=31 align="center" alt="��������" onClick="window.open('Refresh/RefreshNews.asp','fs_main')" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
          <%If HaveValueTF = True then%>
		  <td width="9" align="right" valign="middle" class="Gray">|</td>
          <td width="34" align="right" valign="middle"onClick="window.open('Refresh/Mall_Refresh.asp','fs_main')" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut"><div align="left">�̳�</div></td>
		  <%End if%>
          <td align="right" valign="middle"><img alt="���ز˵�" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" id="CancelImg" border="0" src="Images/cmsv31_close.png" onclick="ShrinkFrame();return false;" class="BtnMouseOut">&nbsp;</td>
        </tr>
      </table> </td>
</tr>
</table>
</body>
</html>
