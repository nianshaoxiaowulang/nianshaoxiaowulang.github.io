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
if Not JudgePopedomTF(Session("Name"),"P070804") then Call ReturnError1()
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
Function SearchFileName()
	dim Rs,strTemp
	strTemp = "<option value=''>所有地址</option>"& vbcrlf
	Set Rs = SErver.CreateObject(G_FS_RS)
	Rs.open "Select FileName From [FS_Help] order by FileName",HelpConn,1,1
	do while not Rs.eof
		If Instr(Lcase(strTemp),">"&Lcase(Rs("FileName")&"<"))=0 Then
			strTemp = strTemp & "<option value='"&Rs("FileName")&"'>"&Rs("FileName")&"</option>"& vbcrlf
		End If
	Rs.movenext
	loop
	Rs.close
	SEt Rs = Nothing
	SearchFileName = "<Select name=FileName>"&strTemp&"</select>"
End Function
Dim GetFileName
GetFileName = SearchFileName
Set Conn = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../CSS/FS_css.css" rel="stylesheet" type="text/css">
<style>td{font-size:12px;line-height:23px;}</style>
<title>检索帮助信息</title>
<script language="javascript">
<!--
function ChkSubmit(obj){
	var url = 'SearchManage.asp?FileName='+obj.FileName.options[obj.FileName.selectedIndex].value+'&PageField='+obj.PageField.value
	window.returnValue = url;
	window.close();
}
-->
</script>
</head>
<body topmargin="0" leftmargin="0" style="margin:0px;padding:0px;">
<table align=center style="background:menu;height:100%;width:100%">
  <tr><td>
	<table cellpadding=0 width="100%" cellspacing=1 align=center style="padding:2px 4px;">
	  <form name=SearchForm onsubmit="return ChkSubmit(this);">
	  <tr>
		<td colspan=2>　<strong>快速检索帮助信息</strong>[注：支持模糊查询]</td>
	  </tr>	  
	  <tr>
		<td width="60" align=center>页面地址</td>
		<td width="*"><%=GetFileName%></td>
	  </tr>
	  <tr>
		<td align=center>关键字</td>
		<td><input type=text name="PageField" value="" size=32></td>
	  </tr>
	  <tr>
		<td colspan=2 align=center>
		<input type=button value=" 确定 " onclick="ChkSubmit(this.form);">　　　
		<input type=button value=" 关闭 " onclick="window.close();">
		</td>
	  </tr>
	  </form>
	</table>
 </td></tr>
</table>
</body>
</html>
<%Set HelpConn = Nothing%>