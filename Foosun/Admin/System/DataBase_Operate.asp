<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
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
dim conn,RsConfig,DBC,SQLStr
set DBC=New DataBaseClass
set conn=DBC.OpenConnection
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not ((JudgePopedomTF(Session("Name"),"P040603")) OR (JudgePopedomTF(Session("Name"),"P040604"))) then Call ReturnError1()
Dim fso,DataBasePath,month1,day1,hour1,minute1,second1,FileName,BackupDatabase
if Request.Form("Action")="CompressDB" then
	if Not JudgePopedomTF(Session("Name"),"P040604") then Call ReturnError1()
	call Reduce()
elseif request.form("Action")="BackUpDB" then
	if Not JudgePopedomTF(Session("Name"),"P040603") then Call ReturnError1()
	Set fso=Server.CreateObject(G_FS_FSO)
	DataBasePath=request.form("DataBasePath")	
	month1=month(now)
	if Month1<10 then Month1="0"&Month1
	day1=day(now)
	if day1<10 then day1="0"&day1
	hour1=hour(now)
	if hour1<10 then hour1="0"&hour1
	minute1=minute(now)
	if minute1<10 then minute1="0"&minute1
	second1=second(now)
	if second1<10 then second1="0"&second1
	if request.form("FileName")="" then
		FileName=Year(now)&Month1&Day1&Hour1&Minute1&Second1
	else
		Filename=request.form("FileName")
	end if
	BackupDatabase=Server.Mappath("../../FooSun_Data")
	if fso.FileExists(Server.Mappath(""&DataBasePath&"")) then
	if fso.FolderExists(BackupDatabase&"\BackupDatabase")=false then fso.CreateFolder(BackupDatabase&"\BackupDatabase")
		fso.CopyFile(Server.Mappath(""&DataBasePath&"")),(server.mappath("../../FooSun_Data\BackupDatabase\Back"&FileName&".mdb"))
		%>
			<script language="javascript">
			alert ("备份成功");
			</script>
		<%
	else
		%>
			<script language="javascript">
			alert ("找不到数据库文件");
			</script>
		<%
	end if
end if
%>
<html>
<head>
<title>数据库备份/压缩</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<body topmargin="2" leftmargin="2"  oncontextmenu="return false;">
<form action="?" method="post" name="DBForm">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="备份数据库" onClick="BackUpDB();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">备份</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35  align="center" alt="压缩数据库" onClick="CompressDB();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">压缩</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
          <td>&nbsp;<input type="hidden" name="Action"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
	<tr> 
      <td height="50" ><div align="center">当前数据库路径：
          <input name="DataBasePath" type="text" size="40" value="../../FooSun_Data/FooSun_Data.mdb">
        </div></td>
    </tr>
</table>
</form>
</body>
</html>
<script language="JavaScript">
function CompressDB()
{
	document.DBForm.Action.value='CompressDB';
	document.DBForm.submit();
}
function BackUpDB()
{
	document.DBForm.Action.value='BackUpDB';
	document.DBForm.submit();
}
</script>
<%
Sub Reduce()
    Dim I
    Dim TargetDB,ResourceDB
    Dim oJetEngine
    Dim Fso
    Const Jet_Conn_Partial = "Provider=Microsoft.Jet.OLEDB.4.0; Data source="
    Set oJetEngine = Server.CreateObject("JRO.JetEngine")
    Set Fso= CreateObject(G_FS_FSO)
    '关闭数据库链接
    Conn.Close
    Set Conn=Nothing
       ResourceDB=Server.MapPath("../../FooSun_Data/FooSun_Data.mdb")

        If Fso.FileExists(ResourceDB) Then
            '建立临时文件
            TargetDB=Server.MapPath("../../FooSun_Data/FooSun_Data.mdb.bak")
            If Fso.FileExists(TargetDB) Then
                Fso.DeleteFile(TargetDB)
            End If
            oJetEngine.CompactDatabase Jet_Conn_Partial&ResourceDB,Jet_Conn_Partial&TargetDB
            Fso.DeleteFile ResourceDB
            Fso.MoveFile TargetDB,ResourceDB
        End If
   
    Set Fso=Nothing
    Set oJetEngine=Nothing
	%>
	<script language="javascript">
	alert ("压缩成功！")
	</script>
	<%
End Sub
%>