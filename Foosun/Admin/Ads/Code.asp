<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="Cls_Ads.asp" -->
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
<!--#include file="../../../Inc/Session.asp" -->

<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P070206") then Call ReturnError()
Conn.Execute("Update FS_Ads set State=0 where State<>0 and (ClickNum>=MaxClick or ShowNum>=MaxShow or ( EndTime<>null and EndTime<="&StrSqlDate&"))")
dim CodeStr,AdsCodeConfig
Set AdsCodeConfig = Conn.Execute("Select DoMain from FS_Config")
CodeStr = AdsCodeConfig("DoMain")&"/JS/AdsJS/"&request("Location")&".js"
%><head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet" type="text/css">
</head>

<title>���������</title>
<body topmargin="0" leftmargin="0">
<table width="75%" border="0" align="center" cellpadding="0" cellspacing="1">
  <tr> 
    <td width="21%" rowspan="3"><div align="center"><img src="../../Images/Info.gif" width="34" height="33"></div></td>
    <td width="79%" height="15">&nbsp;</td>
  </tr>
  <tr> 
    <td>�������ô���Ϊ:</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="2"> <div align="center"> 
        <textarea name="textfield" cols="58" rows="5"><script src=<%=CodeStr%>></script></textarea>
      </div></td>
  </tr>
  <tr> 
    <td colspan="2"> <div align="center"> 
        <input type="button" name="Submit" value=" �� �� " onclick="window.close();">
      </div></td>
  </tr>
  <tr> 
    <td height="10" colspan="2">&nbsp;</td>
  </tr>
</table>
</body>
<script>
  document.all.textfield.select();
</script>
