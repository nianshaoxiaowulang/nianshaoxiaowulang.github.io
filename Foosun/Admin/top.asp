<% Option Explicit %>
<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Const.asp" -->
<%
'==============================================================================
'������ƣ���Ѷ��վ��Ϣ����ϵͳ
'��ǰ�汾��Foosun Content Manager System v3.1 
'���¸��£�2004.12
'==============================================================================
'��ҵע����ϵ��028-85098980-601,602 ����֧�֣�028-85098980-606��607,�ͻ�֧�֣�608
'��Ʒ��ѯQQ��159410,655071,66252421
'����֧��:���г���ʹ�����⣬�����ʵ�bbs.foosun.net���ǽ���ʱ�ش���
'���򿪷�����Ѷ������ & ��Ѷ���������
'��̳֧�֣���Ѷ������̳(http://bbs.foosun.net)
'�ٷ���վ��www.Foosun.net  ��ʾվ�㣺test.cooin.com    
'��վ����ר����www.cooin.com
'==============================================================================
'��Ѱ汾����������ҳ������Ȩ��Ϣ�������ϱ�վLOGO��������
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
<title>�ޱ����ĵ�</title>
<link href="../../Css/Style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../SysJS/PublicJS.js"></script>
</head>

<body onselectstart="return false;" oncontextmenu="return false;">
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr valign="top">
    <td width="267" background="images/cmsv31_02.png"><img src="images/cmsv31_01.png" width="267" height="60" alt=""></td><td width="54" align="right" background="images/cmsv31_02.png">&nbsp;</td><td background="images/cmsv31_02.png"><a href="Menu_Folders.asp?Action=ContentTree" target="nav_folder_area"><img alt="��Ϣ����" src="images/icon_1.png" width="54" border="0" style="cursor:hand"></a><a href="Menu_Folders.asp?Action=Special" target="nav_folder_area"><img alt="ר�����" src="images/icon_2.png" width="54" border="0" style="cursor:hand"></a><a href="Menu_Folders.asp?Action=NetStation" target="nav_folder_area"><img alt="վ�����" src="images/icon_3.png" width="54" border="0" style="cursor:hand"></a><a href="Menu_Folders.asp?Action=System" target="nav_folder_area"><img alt="ϵͳ����" src="images/icon_4.png" width="54" height="60" border="0" style="cursor:hand"></a><a href="Menu_Folders.asp?Action=UpLoad" target="nav_folder_area"><img alt="����Ŀ¼����" src="images/icon_5.png" width="54" height="60" border="0" style="cursor:hand"></a><a href="Menu_Folders.asp?Action=JSManage" target="nav_folder_area"><img alt="JS����" src="images/icon_6.png" width="54" height="60" border="0" style="cursor:hand"></a><a href="Menu_Folders.asp?Action=OrdinaryManage" target="nav_folder_area"><img alt="�������" src="images/icon_7.png" width="54" height="60" border="0" style="cursor:hand"></a><a href="System/ChangePwd.asp" target="fs_main"><img alt="�޸Ĺ���Ա����" src="images/icon_9.png" width="54" height="60" border="0" style="cursor:hand"></a><a href="LoginOut.asp" target="_top"><img alt="�˳�ϵͳ" src="images/icon_a.png" width="54" height="60" border="0" style="cursor:hand"></a> 
    </td>
  </tr>
</table>
</body>
</html>
