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
if Not JudgePopedomTF(Session("Name"),"P040503") then Call ReturnError1()
Dim Path,FileName,EditFile,FileContent,Result
Result = Request.Form("Result")
Path = "../../../Inc"
FileName = "Const.asp"
EditFile = Server.MapPath(Path) & "\" & FileName
Dim FsoObj,FileObj,FileStreamObj
Set FsoObj = Server.CreateObject(G_FS_FSO)
Set FileObj = FsoObj.GetFile(EditFile)
if Result = "" then
	Set FileStreamObj = FileObj.OpenAsTextStream(1)
	if Not FileStreamObj.AtEndOfStream then
		FileContent = FileStreamObj.ReadAll
	else
		FileContent = ""
	end if
else
	Set FileStreamObj = FileObj.OpenAsTextStream(2)
	FileContent = Request.Form("ConstContent")
	FileStreamObj.Write FileContent
	if Err.Number <> 0 then
		%>
		<script language="JavaScript">
			alert('<% = "����ʧ�ܣ��뿽�������´��ļ��ٱ���" %>');window.location='SysConstSet.asp';
		</script>
		<%
		
	else
		%>
		<script language="JavaScript">
			alert('�޸ĳɹ�');window.location='SysConstSet.asp';
		</script>
		<%
	end if
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ϵͳ��������</title>
</head>
<style type="text/css">
<!--
.SysParaButtonStyle {
	border-top-width: 1px;
	border-right-width: 1px;
	border-bottom-width: 1px;
	border-left-width: 1px;
	border-right-style: solid;
	border-bottom-style: solid;
	border-left-style: solid;
	border-right-color: #999999;
	border-bottom-color: #999999;
	border-left-color: #FFFFFF;
	background-color: #E6E6E6;
}
-->
</style>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<body bgcolor="#FFFFFF" topmargin="2" leftmargin="2" scroll=yes  oncontextmenu="return false;">
<form action="" name="Form" method="post">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="����" onClick="document.Form.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td>&nbsp;<input name="Result" type="hidden" id="Result" value="Modify">
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	 <tr> 
      <td colspan="2"><div align="left">
        <font color="#FF0000">&nbsp;&nbsp ע�⣺<br>&nbsp;&nbsp ���õ�ʱ���SysRootDir�����ݿ�·��������·��������У����������������·��������ע�����ã���������ĵ�����ɲ���Ҫ���鷳</font>  
        </div></td>
    </tr>
    <tr> 
      <td colspan="2"><div align="center">
          <textarea name="ConstContent" rows="34" style="width:99%;"><% = FileContent %></textarea>
        </div></td>
    </tr>
    <tr> 
      <td height="30" colspan="2"><div align="center"><font color="#FF0000">ע�⣺�������������ⵥ����&quot;<font color="#0000FF">'</font>&quot;����ȥ�����벻Ҫʹ�ûس���&quot;<font color="#0000FF">&lt;%</font>&quot;��&quot;<font color="#0000FF">%&gt;</font>&quot;����ȥ������һ��ע�⡣����ֻ���ַ�����Ҫ���ӡ�ɾ��</font></div></td>
    </tr>
</table>
</form>
</body>
</html>
<%
Set FsoObj = Nothing
Set FileObj = Nothing
Set FileStreamObj = Nothing
%>
