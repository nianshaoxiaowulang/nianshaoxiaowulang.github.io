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
Dim FuncName,FileName,PageField,HelpContent,HelpSingleContent
Dim HelpID
HelpID = Request.QueryString("ID")

if isNumeric(HelpID)=false or HelpID="" Then
	FuncName = "����İ�����Ϣ"
	FileName = "����İ�����Ϣ"
	PageField = "����İ�����Ϣ"
	HelpContent = "����İ�����Ϣ"
	HelpSingleContent = "����İ�����Ϣ"
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
		FuncName = "����İ�����Ϣ"
		FileName = "����İ�����Ϣ"
		PageField = "����İ�����Ϣ"
		HelpContent = "����İ�����Ϣ"
		HelpSingleContent = "����İ�����Ϣ"
	end if
	tempRs.close
	set tempRs = Nothing
End IF

Set Conn = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>�Ķ������ļ���Ϣ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../CSS/FS_css.css" rel="stylesheet" type="text/css">
<style>td{font-size:12px;line-height:23px;}</style>
<script language="Javascript">
<!--
function zoomimg(img)
{
  //img.style.zoom��ȡimg��������ű�������תΪʮ��������
  var zoom = parseInt(img.style.zoom,10);
  if (isNaN(zoom))
  {    //��zoom������ʱzoomĬ��Ϊ100��
    zoom = 100
  }
  //event.wheelDelta�����ƶ������ƣ�120�����ƣ�120����ʾ����ÿ������10��
  //zoom += event.wheelDelta / 12;
  //��zoom����10��ʱ����������ʾ����
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
    <td width="83" nowrap> <div align="right"><strong>ҳ�湦��</strong></div></td>
    <td width="889" bgcolor="#F7F7F7"><%=FuncName%></td>
  </tr>
  <tr bgcolor="#EFEFEF"> 
    <td nowrap> <div align="right"><strong>ҳ���ַ</strong></div></td>
    <td bgcolor="#F7F7F7"><%=FileName%></td>
  </tr>
  <tr bgcolor="#EFEFEF"> 
    <td nowrap> <div align="right"><strong>�ؼ���</strong></div></td>
    <td bgcolor="#F7F7F7"><%=PageField%></td>
  </tr>
  <tr bgcolor="#EFEFEF"> 
    <td nowrap> <div align="right"><strong>��˵��</strong></div></td>
    <td height="58" valign="top" bgcolor="#F7F7F7"><%=HelpSingleContent%></td>
  </tr>
  <tr bgcolor="#EFEFEF"> 
    <td nowrap> <div align="right"><strong>��ϸ˵��</strong></div></td>
    <td bgcolor="#F7F7F7"><%=HelpContent%></td>
  </tr>
  <tr style="display:none;" bgcolor="#EFEFEF"> 
    <td nowrap bgcolor="#EFEFEF"></td>
    <td bgcolor="#F7F7F7"><a href="addField.asp?ID=<%=HelpID%>" target="_Modify"> �� �� </a></td>
  </tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="40"> 
      <div align="center"><a href="javascript:window.close()"><img src="../Images/Colse.gif" alt="�رմ���" border="0"></a>��<a href="http://help.foosun.net/Search.asp?Keyword=<% = Server.HTMLEncode(Request("HelpKeyWord")) %>&condition=content"; target="_blank"><img src="../Images/ReHelp.gif" width="119" height="28" border="0"></a></div></td>
  </tr>
</table>
</body>
</html>
<%
Set HelpConn = Nothing
%>