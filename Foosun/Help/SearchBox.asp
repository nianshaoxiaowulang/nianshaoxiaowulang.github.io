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
'������ƣ�FoosunHelp System Form FoosunCMS
'��ǰ�汾��Foosun Content Manager System 3.0 ϵ��
'���¸��£�2005.12
'==============================================================================
'��ҵע����ϵ��028-85098980-601,602 ����֧�֣�028-85098980-605��607,�ͻ�֧�֣�608
'��Ʒ��ѯQQ��159410,394226379,125114015,655071
'����֧��:���г���ʹ�����⣬�����ʵ�bbs.foosun.net���ǽ���ʱ�ش���
'���򿪷�����Ѷ������ & ��Ѷ���������
'��̳֧�֣���Ѷ������̳(http://bbs.foosun.net)
'�ٷ���վ��www.Foosun.net  ��ʾվ�㣺test.cooin.com    
'��վ����ר����www.cooin.com
'==============================================================================
'��Ѱ汾����������ҳ������Ȩ��Ϣ�������ϱ�վLOGO��������
'==============================================================================
Function SearchFileName()
	dim Rs,strTemp
	strTemp = "<option value=''>���е�ַ</option>"& vbcrlf
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
<title>����������Ϣ</title>
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
		<td colspan=2>��<strong>���ټ���������Ϣ</strong>[ע��֧��ģ����ѯ]</td>
	  </tr>	  
	  <tr>
		<td width="60" align=center>ҳ���ַ</td>
		<td width="*"><%=GetFileName%></td>
	  </tr>
	  <tr>
		<td align=center>�ؼ���</td>
		<td><input type=text name="PageField" value="" size=32></td>
	  </tr>
	  <tr>
		<td colspan=2 align=center>
		<input type=button value=" ȷ�� " onclick="ChkSubmit(this.form);">������
		<input type=button value=" �ر� " onclick="window.close();">
		</td>
	  </tr>
	  </form>
	</table>
 </td></tr>
</table>
</body>
</html>
<%Set HelpConn = Nothing%>