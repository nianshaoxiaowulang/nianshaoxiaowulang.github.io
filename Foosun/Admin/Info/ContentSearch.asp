<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
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
if Not JudgePopedomTF(Session("Name"),"P010506") then Call ReturnError()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ޱ����ĵ�</title>
</head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<script language="JavaScript">
var ParentLocationStr=dialogArguments.location.href;
function GetSearchKeyWord(LocationStr,SearchStr)
{
	var SearchLocation=LocationStr.lastIndexOf(SearchStr);
	if (SearchLocation!=-1)
	{
		var StartLoc=LocationStr.indexOf('=',SearchLocation);
		var EndLoc=LocationStr.indexOf('&',SearchLocation);
		if (StartLoc!=-1)
		{
			if (EndLoc!=-1)	return LocationStr.slice(StartLoc+1,EndLoc);
			else return LocationStr.slice(StartLoc+1);
		}
		else return '';
	}
	else return '';
}
var DocumentReadyTF=false;
function document.onreadystatechange()
{
	if (DocumentReadyTF) return;
	SetValue();
	DocumentReadyTF=true;
}
function SetValue()
{
	var i=0;
	document.FunctionForm.SearchContent.value=GetSearchKeyWord(ParentLocationStr,'SearchContent');
	document.FunctionForm.SearchBeginTime.value=GetSearchKeyWord(ParentLocationStr,'SearchBeginTime');
	document.FunctionForm.SearchEndTime.value=GetSearchKeyWord(ParentLocationStr,'SearchEndTime');
	for (i=0;i<document.FunctionForm.SearchScope.options.length;i++) if (document.FunctionForm.SearchScope.options(i).value==GetSearchKeyWord(ParentLocationStr,'SearchScope')) document.FunctionForm.SearchScope.options(i).selected=true;
	for (i=0;i<document.FunctionForm.SearchType.options.length;i++) if (document.FunctionForm.SearchType.options(i).value==GetSearchKeyWord(ParentLocationStr,'SearchType')) document.FunctionForm.SearchType.options(i).selected=true;
}
</script>
<body topmargin="0" leftmargin="0">
<div align="center">
  <table width="90%" border="0" cellpadding="0" cellspacing="0">
    <form action="" method="post" name="FunctionForm">
      <tr> 
        <td width="80">����Ŀ�� </td>
        <td height="30"><select style="width:100%;" name="SearchScope">
            <option value="All">ȫ��</option>
            <option value="News">����</option>
            <option value="DownLoad">����</option>
          </select></td>
      </tr>
      <tr> 
        <td>��������</td>
        <td height="30"><select style="width:100%;" name="SearchType">
            <option value="Title">����</option>
            <option value="Content">����</option>
            <option value="KeyWords">�ؼ���</option>
          </select></td>
      </tr>
      <tr> 
        <td>��������</td>
        <td height="30"><input style="width:100%;" type="text" value="" name="SearchContent"></td>
      </tr>
      <tr> 
        <td>��ʼ����</td>
        <td height="30"><input style="width:100%;" onFocus="OpenWindowAndSetValue('../../FunPages/SelectDate.asp',280,120,window,this);" name="SearchBeginTime" value="" readonly type="text" size="19" maxlength="20">
        </td>
      </tr>
      <tr> 
        <td>����ʱ��</td>
        <td height="30"> 
          <input style="width:100%;" onFocus="OpenWindowAndSetValue('../../FunPages/SelectDate.asp',280,120,window,this);" name="SearchEndTime" value="" readonly type="text" size="19" maxlength="20">
        </td>
      </tr>
      <tr> 
        <td height="30" colspan="2"><div align="center"> 
            <input name="Submit" type="button" class="SearchBtnStyle" onClick="dialogArguments.SearchSubmit(document.FunctionForm);window.close();" value=" ȷ �� ">
          </div></td>
      </tr>
    </form>
  </table>
</div>
</body>
</html>
<%
Set Conn = Nothing
%>