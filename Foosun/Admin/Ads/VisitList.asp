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
dim ATypes,ALocation,VFlag,ViSql,ViObj,FileNumber
ATypes = Request("Types")
ALocation = Request("Location")
if ALocation<>"" and isnull(ALocation)=false then
	ALocation = clng(ALocation)
end if
if ATypes = "Shows" then
	if Not JudgePopedomTF(Session("Name"),"P070207") then Call ReturnError()
	VFlag = "2"
elseif ATypes = "Clicks" then
	if Not JudgePopedomTF(Session("Name"),"P070208") then Call ReturnError()
	VFlag = "1"
else
	if Not (JudgePopedomTF(Session("Name"),"P070207") OR JudgePopedomTF(Session("Name"),"P070208")) then Call ReturnError()
	VFlag = "0"
end if
ViSql = "Select * from FS_AdsVisitList where AdsLocation=" & ALocation & " and VisitType=" & VFlag & " order by ID desc"
Set ViObj = Conn.Execute(ViSql)
%>
<html>
<head>
<style type="text/css">
<!--
 BODY   {border: 0; margin: 0; background: buttonface; cursor: default; font-family:����; font-size:9pt;}
 BUTTON {width:5em}
 TABLE  {font-family:����; font-size:9pt}
 P      {text-align:center}
.TempletItem {
	cursor: default;
}
.TempletSelectItem {
	background-color:highlight;
	cursor: default;
	color: white;
}
.ButtonList {
	background-color: buttonface;
	border-top-width: 1px;
	border-right-width: 1px;
	border-bottom-width: 1px;
	border-left-width: 1px;
	border-top-style: solid;
	border-right-style: solid;
	border-bottom-style: solid;
	border-left-style: solid;
	border-top-color: #FFFFFF;
	border-right-color: #999999;
	border-bottom-color: #999999;
	border-left-color: #FFFFFF;
	cursor: default;
	color: red;

}
-->
</style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����ͳ��</title>
</head>
<body leftmargin="2" topmargin="2">
<table width="100%" border="0" cellpadding="0" cellspacing="1">
  <tr> 
    <td width="16%" height="26" class="ButtonList">
<div align="center">����ʱ��</div></td>
    <td width="13%" class="ButtonList"><div align="center">����IP</div></td>
  </tr>
  <%
	  if ViObj.eof then
	  FileNumber = 1
  %>
  <tr> 
    <td colspan="2"><div align="center"><font color="#FF0000">�˹����ʱû�з��ʼ�¼</font></div></td>
  </tr>
  <%
      end if
	  FileNumber = 1
	 do while not ViObj.eof 
  %>
  <tr>  
    <td><div align="center"><font color=blue><%=ViObj("VisitTime")%></font></div></td>
    <td><div align="center"><font color=blue><%=ViObj("VisitIP")%></font></div></td>
  </tr>
	<%
	 ViObj.movenext
	 FileNumber = FileNumber + 1
	 loop
	%></table>
</body>
</html>
<script>
var FileNumber=<% = FileNumber %>;
window.onload=SetWindowHeight;
function SetWindowHeight()
{
	var FileListHeight='';
	if (FileNumber>10)
	{
		FileListHeight=new String(200);
		window.parent.dialogHeight=FileListHeight+'pt';
		document.body.scroll='yes';
	}
	else
	{
		if (FileNumber<3)
		{
			FileListHeight=new String(3*20);
			window.parent.dialogHeight=FileListHeight+'pt';
			document.body.scroll='no';
		}
		else
		{
			FileListHeight=new String(FileNumber*20);
			window.parent.dialogHeight=FileListHeight+'pt';
			document.body.scroll='no';
		}
	}
}
</script>