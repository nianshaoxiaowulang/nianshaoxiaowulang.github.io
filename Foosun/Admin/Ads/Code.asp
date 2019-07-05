<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="Cls_Ads.asp" -->
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
if Not JudgePopedomTF(Session("Name"),"P070206") then Call ReturnError()
Conn.Execute("Update FS_Ads set State=0 where State<>0 and (ClickNum>=MaxClick or ShowNum>=MaxShow or ( EndTime<>null and EndTime<="&StrSqlDate&"))")
dim CodeStr,AdsCodeConfig
Set AdsCodeConfig = Conn.Execute("Select DoMain from FS_Config")
CodeStr = AdsCodeConfig("DoMain")&"/JS/AdsJS/"&request("Location")&".js"
%><head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet" type="text/css">
</head>

<title>广告代码调用</title>
<body topmargin="0" leftmargin="0">
<table width="75%" border="0" align="center" cellpadding="0" cellspacing="1">
  <tr> 
    <td width="21%" rowspan="3"><div align="center"><img src="../../Images/Info.gif" width="34" height="33"></div></td>
    <td width="79%" height="15">&nbsp;</td>
  </tr>
  <tr> 
    <td>本广告调用代码为:</td>
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
        <input type="button" name="Submit" value=" 关 闭 " onclick="window.close();">
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
