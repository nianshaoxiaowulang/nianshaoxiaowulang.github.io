<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
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
			alert ("���ݳɹ�");
			</script>
		<%
	else
		%>
			<script language="javascript">
			alert ("�Ҳ������ݿ��ļ�");
			</script>
		<%
	end if
end if
%>
<html>
<head>
<title>���ݿⱸ��/ѹ��</title>
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
          <td width=35 align="center" alt="�������ݿ�" onClick="BackUpDB();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35  align="center" alt="ѹ�����ݿ�" onClick="CompressDB();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">ѹ��</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
          <td>&nbsp;<input type="hidden" name="Action"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
	<tr> 
      <td height="50" ><div align="center">��ǰ���ݿ�·����
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
    '�ر����ݿ�����
    Conn.Close
    Set Conn=Nothing
       ResourceDB=Server.MapPath("../../FooSun_Data/FooSun_Data.mdb")

        If Fso.FileExists(ResourceDB) Then
            '������ʱ�ļ�
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
	alert ("ѹ���ɹ���")
	</script>
	<%
End Sub
%>