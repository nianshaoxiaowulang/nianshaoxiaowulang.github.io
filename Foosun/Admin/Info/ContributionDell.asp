<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
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
if Not JudgePopedomTF(Session("Name"),"P010604") then Call ReturnError()
    Dim NewsID,NewsObj
	If Request("NewsID")<>"" then
		NewsID = Request("NewsID")
	Else
	   Response.Write("<script>alert(""�������ݴ���"");dialogArguments.location.reload();window.close();</script>")
	   Response.End
	End If
%>
<html>
<head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>���ɾ��</title>
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
      <td>��ȷ��Ҫɾ�����?</td>
    </tr>
  <tr>
    <td height="2">&nbsp;</td>
    <td height="2">&nbsp;</td>
    </tr>
  <tr> 
    <td>&nbsp;</td>
    <td colspan="2"><div align="center"> 
        <input type="submit" name="Submit" value=" ȷ �� ">
        <input type="hidden" name="action" value="trues">
        <input type="button" name="Submit2" value=" ȡ �� " onClick="window.close();">
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