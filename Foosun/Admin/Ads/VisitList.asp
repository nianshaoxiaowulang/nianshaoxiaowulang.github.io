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
 BODY   {border: 0; margin: 0; background: buttonface; cursor: default; font-family:宋体; font-size:9pt;}
 BUTTON {width:5em}
 TABLE  {font-family:宋体; font-size:9pt}
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
<title>访问统计</title>
</head>
<body leftmargin="2" topmargin="2">
<table width="100%" border="0" cellpadding="0" cellspacing="1">
  <tr> 
    <td width="16%" height="26" class="ButtonList">
<div align="center">访问时间</div></td>
    <td width="13%" class="ButtonList"><div align="center">访问IP</div></td>
  </tr>
  <%
	  if ViObj.eof then
	  FileNumber = 1
  %>
  <tr> 
    <td colspan="2"><div align="center"><font color="#FF0000">此广告暂时没有访问记录</font></div></td>
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