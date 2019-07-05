<% Option Explicit %>
<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Const.asp" -->
<%
Dim DBC,Conn,HelpConn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
DBC.ConnStr = "DBQ=" + Server.MapPath("Foosun_help.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
Set HelpConn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../Inc/Session.asp" -->
<!--#include file="../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P070805") then Call ReturnError1()
'==============================================================================
'软件名称：FoosunHelp System Form FoosunCMS
'当前版本：Foosun Content Manager System 3.0 系列
'最新更新：2005.12
'==============================================================================
'商业注册联系：028-85098980-601,602 技术支持：028-85098980-605、607,客户支持：608
'产品咨询QQ：159410,394226379,125114015,655071
'技术支持:所有程序使用问题，请提问到bbs.foosun.net我们将及时回答您
'程序开发：风讯开发组 & 风讯插件开发组
'论坛支持：风讯在线论坛(http://bbs.foosun.net)
'官方网站：www.Foosun.net  演示站点：test.cooin.com    
'网站建设专区：www.cooin.com
'==============================================================================
'免费版本请在新闻首页保留版权信息，并做上本站LOGO友情连接
'==============================================================================
Dim FuncName,FileName,PageField,HelpContent,HelpSingleContent
Dim HelpID
HelpID = Request.QueryString("ID")

if isNumeric(HelpID)=false or HelpID="" Then
	FuncName = "错误的帮助信息"
	FileName = "错误的帮助信息"
	PageField = "错误的帮助信息"
	HelpContent = "错误的帮助信息"
	HelpSingleContent = "错误的帮助信息"
Else
	Dim tempRs
	Set tempRs = Server.CreateObject(G_FS_RS)
	tempRs.open "Select * From [Fs_Help] where id="&Clng(HelpID),HelpConn,1,1
	if not tempRs.eof then
		FuncName = tempRs("FuncName")
		FileName = tempRs("FileName")
		PageField = tempRs("PageField")
		HelpContent = Replace(tempRs("HelpContent"),"../../Files/","../../"&UpFiles&"/")
		HelpSingleContent = Replace(tempRs("HelpSingleContent"),"../../Files/","../../"&UpFiles&"/")
	Else
		FuncName = "错误的帮助信息"
		FileName = "错误的帮助信息"
		PageField = "错误的帮助信息"
		HelpContent = "错误的帮助信息"
		HelpSingleContent = "错误的帮助信息"
	end if
	tempRs.close
	set tempRs = Nothing
End IF

Set Conn = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>阅读帮助文件信息</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../CSS/FS_css.css" rel="stylesheet" type="text/css">
<style>td{font-size:12px;line-height:23px;}</style>
<script language="Javascript">
<!--
function zoomimg(img)
{
  //img.style.zoom获取img对象的缩放比例，并转为十进制整数
  var zoom = parseInt(img.style.zoom,10);
  if (isNaN(zoom))
  {    //当zoom非数字时zoom默认为100％
    zoom = 100
  }
  //event.wheelDelta滚轮移动量上移＋120，下移－120；显示比例每次增减10％
  //zoom += event.wheelDelta / 12;
  //当zoom大于10％时重新设置显示比例
  if (zoom == 100)
  {
  	if(img.alt == "" )
	{
		img.style.zoom = 25 + "%";
	}
	else
  		img.style.zoom = img.alt + "%";
  }
  else
  	img.style.zoom = 100 + "%";	
}
-->
</script>
</head>

<body topmargin="4" leftmargin="2">
<table cellpadding=4 width="98%" cellspacing=1 align=center bgcolor="#DEDEDE" style="padding:0px 4px;">
  <tr bgcolor="#EFEFEF"> 
    <td width="83" nowrap> <div align="right"><strong>页面功能</strong></div></td>
    <td width="889" bgcolor="#F7F7F7"><%=FuncName%></td>
  </tr>
  <tr bgcolor="#EFEFEF"> 
    <td nowrap> <div align="right"><strong>页面地址</strong></div></td>
    <td bgcolor="#F7F7F7"><%=FileName%></td>
  </tr>
  <tr bgcolor="#EFEFEF"> 
    <td nowrap> <div align="right"><strong>关键字</strong></div></td>
    <td bgcolor="#F7F7F7"><%=PageField%></td>
  </tr>
  <tr bgcolor="#EFEFEF"> 
    <td nowrap> <div align="right"><strong>简单说明</strong></div></td>
    <td height="58" valign="top" bgcolor="#F7F7F7"><%=HelpSingleContent%></td>
  </tr>
  <tr bgcolor="#EFEFEF"> 
    <td nowrap> <div align="right"><strong>详细说明</strong></div></td>
    <td bgcolor="#F7F7F7"><%=HelpContent%></td>
  </tr>
  <tr style="display:none;" bgcolor="#EFEFEF"> 
    <td nowrap bgcolor="#EFEFEF"></td>
    <td bgcolor="#F7F7F7"><a href="addField.asp?ID=<%=HelpID%>" target="_Modify"> 修 改 </a></td>
  </tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="40"> 
      <div align="center"><a href="javascript:window.close()"><img src="../Images/Colse.gif" alt="关闭窗口" border="0"></a>　<a href="http://help.foosun.net/Search.asp?Keyword=<% = Server.HTMLEncode(Request("HelpKeyWord")) %>&condition=content"; target="_blank"><img src="../Images/ReHelp.gif" width="119" height="28" border="0"></a></div></td>
  </tr>
</table>
</body>
</html>
<%
Set HelpConn = Nothing
%>