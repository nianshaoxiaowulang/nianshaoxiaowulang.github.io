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
if Not JudgePopedomTF(Session("Name"),"P040605") then Call ReturnError1()
%>
<html>
<head>
<title>ִ��SQL�ű�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<style type="text/css">
<!--
.SysParaButtonStyle {
	border-top-width: 1px;
	border-right-width: 1px;
	border-bottom-width: 1px;
	border-left-width: 1px;
	border-right-style: solid;
	border-bottom-style: solid;
	border-left-style: solid;
	border-right-color: #999999;
	border-bottom-color: #999999;
	border-left-color: #FFFFFF;
	background-color: #E6E6E6;
}
-->
</style>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<body topmargin="2" leftmargin="2" scroll="no" oncontextmenu="return false;">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
		  <td width=35 align="center" alt="ִ��SQL���" onClick="ExecuteSql();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">ִ��</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td>&nbsp;<input type=hidden name=operation value=Modify></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="80%"><textarea name="Content" rows="5" wrap="OFF" style="width:100%;"></textarea></td>
  </tr>
  <tr> 
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><div align="center"><font color="red">ע��һ��ֻ��ִ��һ��Sql��䡣������SQL����Ϥ���뾡����Ҫʹ�á�����һ���������������ġ�</font></div></td>
        </tr>
	  </table></td>
  </tr>
  <tr> 
    <td><iframe id="ResultShowFrame" scrolling="yes" src="DataBase_SqlResult.asp" style="width:100%;" frameborder=1></iframe></td>
  </tr>
</table>
</body>
</html>
<%
Set Conn=nothing
%>
<script language="JavaScript">
function ExecuteSql()
{
	var FormObj=frames["ResultShowFrame"].document.ExecuteForm;
	if (document.all.Content.value!='')
	{
		FormObj.Sql.value=document.all.Content.value;
		FormObj.Result.value='Submit';
		FormObj.submit();
		FormObj.Result.value='';
	}
	else alert('����дSQL���');
}
function SetEditAreaHeight()
{
	var BodyHeight=document.body.clientHeight;
	var EditAreaHeight=BodyHeight-159;
	document.all.ResultShowFrame.height=EditAreaHeight;
}
SetEditAreaHeight();
window.onresize=SetEditAreaHeight;
</script>